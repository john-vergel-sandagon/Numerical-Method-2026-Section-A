"""
Series 1 - Exercise 1
=====================
Laboratory exercise statement:

    Do the Python implementation up to nanosecond of

        (1 + 1/n)^n

    for how-often periods: yearly, twice a year, quarterly, monthly,
    weekly, daily, hourly, every minute, every second, every millisecond,
    every microsecond, every nanosecond.

    Draw a Histogram in Matplotlib.

Theory
------
The number e = 2.718281828... is defined as the limit

        e = lim_{n -> inf} (1 + 1/n)^n

which is exactly the "continuously compounded interest" limit: if you
compound 100% annual interest n times per year, the accumulated factor
after one year is (1 + 1/n)^n, and it converges to e as the compounding
period shrinks.

The classical first order error estimate (from the Taylor expansion of
n * ln(1 + 1/n) = 1 - 1/(2n) + 1/(3n^2) - ...) gives

        e - (1 + 1/n)^n  ~  e / (2n)      for large n

so the error should roughly halve every time n doubles.

Numerical note
--------------
For very large n (every microsecond / every nanosecond in a year, n is
of the order of 1e13 - 1e16) the naive expression `1 + 1/n` loses all of
its precision in double precision floating point (1/n becomes smaller
than machine epsilon relative to 1, so `1 + 1/n` rounds to exactly 1.0).
To keep the computation accurate we instead use

        (1 + 1/n)^n = exp( n * log1p(1/n) )

`math.log1p(x)` computes log(1 + x) accurately even when x is tiny,
which avoids the cancellation problem above.
"""
import math
import matplotlib.pyplot as plt

# --- how-often periods, expressed as "times compounded per year" -------
labels = [
    "yearly", "twice a year", "quarterly", "monthly", "weekly", "daily",
    "hourly", "every minute", "every second", "every millisecond",
    "every microsecond", "every nanosecond",
]

SECONDS_PER_YEAR = 365 * 24 * 60 * 60

n_values = [
    1,                                  # yearly
    2,                                  # twice a year
    4,                                  # quarterly
    12,                                 # monthly
    52,                                 # weekly
    365,                                # daily
    365 * 24,                          # hourly
    365 * 24 * 60,                     # every minute
    SECONDS_PER_YEAR,                  # every second
    SECONDS_PER_YEAR * 1_000,          # every millisecond
    SECONDS_PER_YEAR * 1_000_000,      # every microsecond
    SECONDS_PER_YEAR * 1_000_000_000,  # every nanosecond
]


def compound_value(n: int) -> float:
    """Numerically stable (1 + 1/n)^n using log1p, safe even for n ~ 1e16."""
    return math.exp(n * math.log1p(1.0 / n))


values = [compound_value(n) for n in n_values]
e = math.e
errors = [abs(v - e) for v in values]

# --- print the table (as requested in the exercise statement) ----------
print(f"{'How often':<20} {'n':>18} {'(1+1/n)^n':>14} {'error':>14}")
for lbl, n, v, err in zip(labels, n_values, values, errors):
    print(f"{lbl:<20} {n:>18} {v:>14.6f} {err:>14.3e}")
print(f"\ne = {e:.6f}")

# --- plotting ------------------------------------------------------------
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle("Exercise 1: Convergence of (1 + 1/n)^n to e", fontweight="bold")

bars1 = ax1.bar(labels, values, color="royalblue", edgecolor="black")
ax1.axhline(e, color="red", linestyle="--", label=f"e = {e:.6f}")
for b, v in zip(bars1, values):
    ax1.text(b.get_x() + b.get_width() / 2, v, f"{v:.6f}",
              rotation=90, va="bottom", ha="center", fontsize=8)
ax1.set_ylabel("(1 + 1/n)^n")
ax1.set_title("Value per compounding period")
ax1.set_ylim(1.9, 3.05)
plt.setp(ax1.get_xticklabels(), rotation=45, ha="right")
ax1.legend(loc="lower right")

ax2.bar(labels, errors, color="orange", edgecolor="black")
ax2.set_yscale("log")
ax2.set_ylabel("|(1 + 1/n)^n - e|  (log scale)")
ax2.set_title("Error shrinks like e/(2n)")
plt.setp(ax2.get_xticklabels(), rotation=45, ha="right")

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("exercise 1.png", dpi=150)
plt.show()
