"""
Series 1 - Exercise 2
=====================
Laboratory exercise statement:

    Create the Python implementation of the table:

        (a^h - 1) / h

    for three different bases a (a = 2, a = e = 2.71828..., a = 3) as
    h shrinks (h = 0.1, 0.01, 0.001, 0.0001, ...), tolerance 1e-6.

    Draw a Histogram in Matplotlib.

Theory
------
The quantity (a^h - 1) / h is the symmetric difference quotient of the
function f(x) = a^x at x = 0, taken with step h:

        (a^h - a^0) / h  ->  f'(0)   as h -> 0

Since f(x) = a^x = e^(x ln a), we have f'(x) = ln(a) * a^x, so
f'(0) = ln(a). Therefore the difference quotient settles at ln(a) as h
shrinks towards 0 -- this is exactly the definition-of-the-derivative
argument that explains why `e` is the unique base whose exponential is
its own derivative (ln(e) = 1).

We stop shrinking h once two consecutive values differ by less than the
requested tolerance (1e-6), matching the "settles at" row of the table
in the exercise.
"""
import math
import matplotlib.pyplot as plt

TOLERANCE = 1e-6

bases = {
    "a=2": 2.0,
    "a=e": math.e,
    "a=3": 3.0,
}

h_values = [0.1, 0.01, 0.001, 0.0001, 1e-5, 1e-6, 1e-7]


def difference_quotient(a: float, h: float) -> float:
    return (a ** h - 1.0) / h


# --- build the table -----------------------------------------------------
results = {name: [difference_quotient(a, h) for h in h_values] for name, a in bases.items()}

# figure out at which h each base has "settled" within the tolerance of ln(a)
settle_h = {}
for name, a in bases.items():
    target = math.log(a)
    settle_h[name] = next(
        (h for h, v in zip(h_values, results[name]) if abs(v - target) < TOLERANCE),
        None,
    )

# --- print the table -----------------------------------------------------
header = f"{'h':>10}" + "".join(f"{name:>14}" for name in bases)
print(header)
for i, h in enumerate(h_values):
    row = f"{h:>10g}" + "".join(f"{results[name][i]:>14.4f}" for name in bases)
    print(row)
print(f"{'settles at':>10}" + "".join(f"{math.log(a):>14.4f}" for a in bases.values()))
print(f"\nTolerance used: {TOLERANCE:g}")
for name in bases:
    h = settle_h[name]
    print(f"  {name}: settles within tolerance at h = {h}" if h else
          f"  {name}: did not settle within the tested h range")

# --- plotting --------------------------------------------------------------
x = range(len(h_values))
width = 0.25
colors = {"a=2": "royalblue", "a=e": "green", "a=3": "red"}

fig, ax = plt.subplots(figsize=(11, 6))
for i, (name, a) in enumerate(bases.items()):
    offset = (i - 1) * width
    ax.bar(
        [xi + offset for xi in x], results[name], width,
        label=f"{name}  (ln a = {math.log(a):.4f})",
        color=colors[name], edgecolor="black",
    )
    ax.axhline(math.log(a), color=colors[name], linestyle="--", linewidth=1)

ax.set_xticks(list(x))
ax.set_xticklabels([f"h = {h:g}" for h in h_values])
ax.set_ylabel("(a^h - 1)/h")
ax.set_title("Bars: difference quotient per h.  Dashed lines: the limit ln(a).")
fig.suptitle("Exercise 2: (a^h - 1)/h settling to ln(a)", fontweight="bold")
ax.set_ylim(0, 1.35)
ax.legend(loc="upper center", ncol=3)

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("exercise 2.png", dpi=150)
plt.show()
