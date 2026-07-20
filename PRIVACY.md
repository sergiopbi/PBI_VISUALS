# Privacy Policy — Smart Bucket Chart (Power BI custom visual)

**Last updated:** July 2026

Smart Bucket Chart is a Power BI custom visual developed by Sergio ("the developer",
"we"). This policy explains how the visual handles data.

## Data collection

Smart Bucket Chart does **not** collect, store, transmit, or share any data.

- It does not make network requests of any kind (no HTTP/HTTPS calls, no
  WebSockets, no telemetry, no analytics, no crash reporting).
- It does not use browser local storage, cookies, or any other persistent
  storage.
- It does not access any external service or resource.

This is reflected in the visual's declared capabilities: the `privileges`
array in `capabilities.json` is empty, meaning Power BI grants the visual no
special access at all.

## Data processing

All data shown by Smart Bucket Chart comes exclusively from the fields the
report author drags into the visual's data roles (Dimension, Bucket by,
Values, Tooltips) inside their own Power BI report. This data is processed
entirely **locally, in the user's browser or the Power BI Desktop/Service
rendering engine**, for the sole purpose of drawing the chart. Nothing is
copied, cached beyond the current rendering session, or sent anywhere else.

When the report is closed or the visual is removed from the page, no trace
of that data is retained by the visual.

## Third parties

Smart Bucket Chart does not integrate with, or send data to, any third-party
service, library backend, or analytics provider.

## Changes to this policy

If this visual's functionality changes in a way that affects this policy
(for example, if a future version adds an optional network-dependent
feature), this document will be updated accordingly, and the version history
will remain available in the source repository.

## Contact

Questions about this policy or the visual can be raised via the support
channel listed in the AppSource listing, or as an issue on the source
repository:
https://github.com/sergiopbi/PBI_VISUALS
