from __future__ import annotations
import sys

class ProgressReporter:
    """Affiche la progression sur stderr (ligne courante / total)."""

    def __init__(
        self,
        label: str,
        total: int,
        *,
        enabled: bool = True,
        interval: int = 500,
    ) -> None:
        self.label = label
        self.total = max(total, 1)
        self.enabled = enabled
        self.interval = max(interval, 1)
        self._last_shown = 0

    def update(self, current: int) -> None:
        if not self.enabled or current < 1:
            return
        if current != self.total and current % self.interval != 0 and current != 1:
            return
        self._last_shown = current
        percent = int(current * 100 / self.total)
        line = f"      {self.label} : {current} / {self.total} ({percent}%)"
        sys.stderr.write(f"\r{line:<70}")
        sys.stderr.flush()

    def finish(self, detail: str = "") -> None:
        if not self.enabled:
            return
        if self._last_shown != self.total:
            sys.stderr.write(f"\r      {self.label} : {self.total} / {self.total}   ")
        sys.stderr.write("\n")
        if detail:
            sys.stderr.write(f"      {detail}\n")
        sys.stderr.flush()




