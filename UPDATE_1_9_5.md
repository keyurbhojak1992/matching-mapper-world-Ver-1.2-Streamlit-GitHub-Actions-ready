# Update 1.9.5

Auto Pricing production cleanup:
- Removes only fully identical output rows after merge/window processing.
- Preserves repeated activity whenever any output field differs.
- Keeps automatic current-date processing for live runs.
- Keeps full boundary-date replacement, saved removed-party filtering, Sales SPer/SRate refresh, and blank Lab/Flo -> NON.
