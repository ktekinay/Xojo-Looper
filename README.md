# Xojo-Looper
Process stuff in the background will full UI access

## What It Does

A Xojo thread lets you do stuff in the background but does not let you access the UI directly. This class emulates a Thread by doing its own time-slicing around iterations. Since it runs in the Main Thread, you will have full access to your responsive UI.

See the Harness project for examples.

## Contributions

Contributions to this project are welcome. Fork it to your own repo, then submit changes. However, be forewarned that only those changes that we deem universally useful will be included.

## Who Did This?

This project was designed and implemented by:

* Kem Tekinay (ktekinay at mactechnologies.com)

## Release Notes

1.0 (Sept. 6, 2019)

* Initial release.
