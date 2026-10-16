"""Run `py test_scorecard.py` to confirm your College Scorecard key works."""
import os, tomllib
import scorecard

key = None
for path in [".streamlit/secrets.toml", "secrets.toml"]:
    try:
        with open(path, "rb") as f:
            key = tomllib.load(f).get("SCORECARD_API_KEY")
            if key:
                break
    except FileNotFoundError:
        pass
key = key or os.environ.get("SCORECARD_API_KEY")
if not key:
    raise SystemExit("No SCORECARD_API_KEY found. Put it in secrets.toml (free key: https://api.data.gov/signup/)")

for name in ["Massachusetts Institute of Technology", "University of California, Berkeley", "University of Texas at Austin", "Texas A&M University"]:
    info = scorecard.lookup_school(name, key, "users.db")
    if not info:
        print(f"{name}: NOT FOUND")
        continue
    print(f"{name} -> {info['name']}")
    print({k: info[k] for k in ("admit_rate", "sat_25", "sat_75", "net_price", "grad_rate", "size")})
    print("   est. chance for a 1450 SAT:", scorecard.estimate_chance(info, sat=1450))
