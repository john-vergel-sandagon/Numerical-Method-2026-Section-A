"""
Series 1 - Exercise 3
=====================
Laboratory exercise statement:

    Do the Python implementation of the following:

        e^x = sum_{n=0}^{inf} x^n / n!

    Up to 10,000 (terms), for x = 1.

    Draw a Histogram in Matplotlib.

Theory
------
The Maclaurin (Taylor at 0) series expansion of e^x is

        e^x = 1 + x + x^2/2! + x^3/3! + ... = sum_{n=0}^{inf} x^n/n!

Evaluated at x = 1 this gives the classic series for Euler's number:

        e = sum_{n=0}^{inf} 1/n!

We compute the partial sums S_N = sum_{n=0}^{N} 1/n! for a selection of
N values up to 10,000 and compare them to math.e.

Because 1/n! shrinks *factorially* fast, the series converges far faster
than the (1 + 1/n)^n limit of Exercise 1: by N = 20 the terms are already
below double-precision machine epsilon (~2.22e-16), so the partial sum
stops changing -- adding more terms (up to 10,000) cannot improve the
result any further. This is illustrated on the right-hand plot, which
shows how many correct decimal digits each partial sum buys.
"""
import math
import matplotlib.pyplot as plt

MAX_N = 10_000
x = 1.0
e = math.e

# --- build partial sums S_N incrementally up to N = 10,000 --------------
term = 1.0          # x^0 / 0!
partial_sum = 1.0
partial_sums = {0: partial_sum}
for n in range(1, MAX_N + 1):
    term *= x / n    # turns term_{n-1} into term_n = x^n / n!
    partial_sum += term
    partial_sums[n] = partial_sum

# N values shown on the histogram (as in the exercise's example table)
N_display = [1, 2, 3, 5, 10, 15, 20, 30, 50, 100, 1_000, 10_000]
S_values = [partial_sums[n] for n in N_display]
errors = [abs(s - e) for s in S_values]

# --- print the table ------------------------------------------------------
print(f"{'N':>8} {'S_N':>16} {'|S_N - e|':>16}")
for n, s, err in zip(N_display, S_values, errors):
    print(f"{n:>8} {s:>16.10f} {err:>16.3e}")
print(f"\ne = {e:.10f}")

# --- correct decimal digits per N (for the accuracy-band plot) -----------
def correct_digits(err: float) -> float:
    if err <= 0:
        return 16.0  # cap at ~double precision
    return -math.log10(err)


digits = [correct_digits(err) for err in errors]


def band_color(err: float) -> str:
    if err > 1e-2:
        return "red"       # rough
    if err > 1e-6:
        return "orange"    # engineering
    if err > 1e-14:
        return "green"     # high precision
    return "royalblue"     # machine precision


colors = [band_color(err) for err in errors]

# --- plotting ---------------------------------------------------------------
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle("Exercise 3: e^x = sum x^n / n!  (x = 1), summation up to N = 10,000 terms",
             fontweight="bold")

# Left: histogram of partial sums
bars = ax1.bar([str(n) for n in N_display], S_values, color="purple", edgecolor="black")
ax1.axhline(e, color="red", linestyle="--", label=f"e = {e:.6f}")
for b, s in zip(bars, S_values):
    ax1.text(b.get_x() + b.get_width() / 2, s, f"{s:.6f}",
              rotation=90, va="bottom", ha="center", fontsize=8)
ax1.set_xlabel("number of terms N in the summation")
ax1.set_ylabel("S_N = sum_(n=0..N) 1/n!")
ax1.set_title("Histogram of the partial sums")
ax1.set_ylim(0.9, 3.05)
plt.setp(ax1.get_xticklabels(), rotation=45, ha="right")
ax1.legend(loc="lower right")

# Right: log-log accuracy plot with color-coded accuracy bands
ax2.plot(N_display, errors, color="gray", linewidth=1, zorder=1)
ax2.scatter(N_display, errors, c=colors, s=80, edgecolor="black", zorder=2)
for n, err, d in zip(N_display, errors, digits):
    if n == 20:
        label = "15 digits from here on"
    elif n > 20:
        label = None  # avoid repeating/overlapping the same label
    else:
        label = f"{d:.0f} digits"
    if label:
        ax2.annotate(label, (n, err), textcoords="offset points",
                     xytext=(0, 8), fontsize=8, ha="center",
                     color=band_color(err), fontweight="bold")
for n in N_display:
    ax2.axvline(n, color="lightgray", linewidth=0.5)
ax2.axhspan(0, 1e-14, color="royalblue", alpha=0.1)
ax2.text(N_display[0], 3e-15, "machine precision zone: from N = 20 on, adding "
         "terms cannot improve a float sum", fontsize=7, color="royalblue")
ax2.set_xscale("log")
ax2.set_yscale("log")
ax2.set_xlabel("number of terms N in the summation  (log scale)")
ax2.set_ylabel("|S_N - e|  (log scale)")
ax2.set_title("How many correct digits each N buys (log-log)")

legend_handles = [
    plt.Line2D([0], [0], marker="o", color="w", markerfacecolor="red", markersize=10, label="rough (error > 1e-2)"),
    plt.Line2D([0], [0], marker="o", color="w", markerfacecolor="orange", markersize=10, label="engineering (1e-6 to 1e-2)"),
    plt.Line2D([0], [0], marker="o", color="w", markerfacecolor="green", markersize=10, label="high precision (< 1e-6)"),
    plt.Line2D([0], [0], marker="o", color="w", markerfacecolor="royalblue", markersize=10, label="machine precision (~2e-16)"),
]
ax2.legend(handles=legend_handles, title="accuracy band", loc="upper right", fontsize=8)

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("exercise 3.png", dpi=150)
plt.show()
