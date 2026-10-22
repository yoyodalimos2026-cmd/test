import subprocess
import os
import sys
import tempfile
import unittest
from pathlib import Path

from app.main_window import GitPushTask


def git(repository: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=repository,
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


class GitPushTaskTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary_directory.cleanup)
        self.root = Path(self.temporary_directory.name)
        self.repository = self.root / "working"
        self.remote = self.root / "remote.git"
        git(self.root, "init", "--bare", str(self.remote))
        git(self.root, "init", "--initial-branch=main", str(self.repository))
        git(self.repository, "config", "user.name", "Test User")
        git(self.repository, "config", "user.email", "test@example.com")
        (self.repository / "source.py").write_text("version = 1\n", encoding="utf-8")
        (self.repository / "other.py").write_text("unchanged = 1\n", encoding="utf-8")
        git(self.repository, "add", ".")
        git(self.repository, "commit", "-m", "Initial commit")

    def _task(
        self,
        branch: str,
        paths: list[str],
        remote: Path | None = None,
    ) -> tuple[GitPushTask, list[str], list[str]]:
        task = GitPushTask(
            str(remote or self.remote),
            branch,
            "Update selected source",
            paths,
            repo_path=self.repository,
        )
        finished: list[str] = []
        failed: list[str] = []
        task.signals.finished.connect(lambda: finished.append("done"))
        task.signals.failed.connect(failed.append)
        task.run()
        return task, finished, failed

    def test_pushes_only_selected_file_to_main(self) -> None:
        (self.repository / "source.py").write_text("version = 2\n", encoding="utf-8")
        (self.repository / "other.py").write_text("changed = 2\n", encoding="utf-8")

        _, finished, failed = self._task("main", ["source.py"])

        self.assertEqual(finished, ["done"])
        self.assertEqual(failed, [])
        self.assertEqual(
            git(self.root, "--git-dir", str(self.remote), "show", "main:source.py"),
            "version = 2",
        )
        self.assertEqual(
            git(self.root, "--git-dir", str(self.remote), "show", "main:other.py"),
            "unchanged = 1",
        )
        self.assertIn("other.py", git(self.repository, "status", "--short"))

    def test_creates_and_pushes_new_branch(self) -> None:
        (self.repository / "source.py").write_text("version = 3\n", encoding="utf-8")

        _, finished, failed = self._task("feature/test", ["source.py"])

        self.assertEqual(finished, ["done"])
        self.assertEqual(failed, [])
        self.assertEqual(git(self.repository, "branch", "--show-current"), "feature/test")
        self.assertEqual(
            git(self.root, "--git-dir", str(self.remote), "show", "feature/test:source.py"),
            "version = 3",
        )

    def test_checks_out_and_pushes_to_existing_remote_branch(self) -> None:
        git(self.repository, "remote", "add", "origin", str(self.remote))
        git(self.repository, "switch", "-c", "feature/existing")
        git(self.repository, "push", "-u", "origin", "feature/existing")
        git(self.repository, "switch", "main")
        git(self.repository, "branch", "-D", "feature/existing")
        git(self.repository, "fetch", "origin")
        (self.repository / "source.py").write_text("version = 8\n", encoding="utf-8")

        _, finished, failed = self._task("feature/existing", ["source.py"])

        self.assertEqual(finished, ["done"])
        self.assertEqual(failed, [])
        self.assertEqual(git(self.repository, "branch", "--show-current"), "feature/existing")
        self.assertEqual(
            git(self.root, "--git-dir", str(self.remote), "show", "feature/existing:source.py"),
            "version = 8",
        )

    def test_retries_a_failed_push_without_creating_an_empty_commit(self) -> None:
        (self.repository / "source.py").write_text("version = 5\n", encoding="utf-8")
        unavailable_remote = self.root / "not-created-yet.git"

        _, finished, failed = self._task("feature/retry", ["source.py"], unavailable_remote)

        self.assertEqual(finished, [])
        self.assertEqual(len(failed), 1)
        git(self.root, "init", "--bare", str(unavailable_remote))

        _, finished, failed = self._task("feature/retry", ["source.py"], unavailable_remote)

        self.assertEqual(finished, ["done"])
        self.assertEqual(failed, [])
        self.assertEqual(git(self.repository, "rev-list", "--count", "feature/retry"), "2")
        self.assertEqual(
            git(self.root, "--git-dir", str(unavailable_remote), "show", "feature/retry:source.py"),
            "version = 5",
        )

    def test_refuses_to_commit_pre_staged_files(self) -> None:
        (self.repository / "source.py").write_text("version = 4\n", encoding="utf-8")
        git(self.repository, "add", "source.py")

        _, finished, failed = self._task("main", ["source.py"])

        self.assertEqual(finished, [])
        self.assertEqual(len(failed), 1)
        self.assertIn("déjà indexés", failed[0])
        self.assertEqual(git(self.repository, "branch", "--show-current"), "main")

    def test_commit_failure_leaves_selected_changes_unstaged_for_retry(self) -> None:
        (self.repository / "source.py").write_text("version = 6\n", encoding="utf-8")
        hook = self.repository / ".git" / "hooks" / "pre-commit"
        hook.write_text("#!/bin/sh\nexit 1\n", encoding="utf-8")
        hook.chmod(0o755)

        _, finished, failed = self._task("main", ["source.py"])

        self.assertEqual(finished, [])
        self.assertEqual(len(failed), 1)
        self.assertEqual(git(self.repository, "diff", "--cached", "--name-only"), "")
        hook.unlink()

        _, finished, failed = self._task("main", ["source.py"])

        self.assertEqual(finished, ["done"])
        self.assertEqual(failed, [])

    def test_askpass_uses_temporary_private_credentials_file(self) -> None:
        task = GitPushTask(
            str(self.remote),
            "main",
            "Update source",
            ["source.py"],
            username="git-user",
            token="temporary-token",
            repo_path=self.repository,
        )

        with task._askpass_environment() as environment:
            credentials_path = Path(environment["GIT_GUI_CREDENTIALS_FILE"])
            helper_path = credentials_path.parent / "askpass.py"
            username = subprocess.run(
                [sys.executable, str(helper_path), "Username for 'https://example.com':"],
                check=True,
                capture_output=True,
                text=True,
                env={**os.environ, **environment},
            )
            token = subprocess.run(
                [sys.executable, str(helper_path), "Password for 'https://example.com':"],
                check=True,
                capture_output=True,
                text=True,
                env={**os.environ, **environment},
            )

            self.assertEqual(username.stdout.strip(), "git-user")
            self.assertEqual(token.stdout.strip(), "temporary-token")
            self.assertEqual(credentials_path.stat().st_mode & 0o777, 0o600)
            self.assertNotIn("temporary-token", environment.values())

        self.assertFalse(credentials_path.exists())

    def test_explicit_https_credentials_are_not_persisted_on_task(self) -> None:
        (self.repository / "source.py").write_text("version = 7\n", encoding="utf-8")
        task = GitPushTask(
            str(self.remote),
            "feature/https-auth",
            "Update source",
            ["source.py"],
            username="git-user",
            token="temporary-token",
            repo_path=self.repository,
        )
        finished: list[str] = []
        failed: list[str] = []
        task.signals.finished.connect(lambda: finished.append("done"))
        task.signals.failed.connect(failed.append)

        task.run()

        self.assertEqual(finished, ["done"])
        self.assertEqual(failed, [])
        self.assertEqual(task.username, "")
        self.assertEqual(task.token, "")
        self.assertEqual(
            git(self.root, "--git-dir", str(self.remote), "show", "feature/https-auth:source.py"),
            "version = 7",
        )


if __name__ == "__main__":
    unittest.main()
