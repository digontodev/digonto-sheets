"""Row / column geometry with sparse custom sizes (so one million rows cost nothing)."""
from bisect import bisect_left


class Axis:
    def __init__(self, n, default):
        self.n, self.d = n, default
        self.custom = {}
        self._keys, self._cum = [], []

    def set(self, custom, hidden=()):
        self.custom = {k: v for k, v in custom.items() if k < self.n}
        for h in hidden:
            if h < self.n:
                self.custom[h] = 0
        self._rebuild()

    def _rebuild(self):
        self._keys = sorted(self.custom)
        acc, cum = 0, []
        for k in self._keys:
            acc += self.custom[k] - self.d
            cum.append(acc)
        self._cum = cum

    def size(self, i):
        return self.custom.get(i, self.d)

    def hidden(self, i):
        return self.custom.get(i) == 0

    def pos(self, i):
        j = bisect_left(self._keys, i)
        return i * self.d + (self._cum[j - 1] if j else 0)

    def total(self):
        return self.pos(self.n)

    def index_at(self, px):
        """index of the row/column containing px (clamped)"""
        if px <= 0:
            return 0
        lo, hi = 0, self.n - 1
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if self.pos(mid) <= px:
                lo = mid
            else:
                hi = mid - 1
        return lo

    def nearest_visible(self, i, step):
        while 0 <= i < self.n and self.hidden(i):
            i += step
        return min(max(i, 0), self.n - 1)
