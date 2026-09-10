# Update 1.9.2 — Auto Pricing Stock Upload Simplification

- Removed the Previous Stock upload from Auto Pricing > Full Pipeline.
- Stock is always rebuilt from the latest Stock RAW snapshot.
- The finalizer no longer reads a previous Stock workbook even for headers; the canonical Stock output schema is built into the processor.
- Previous history uploads remain required for Sales 120, Sales 180, View, Hold, Offer, Wishlist, and Cart.
- Git-default Auto Pricing template and Removed Parties behavior is unchanged.
