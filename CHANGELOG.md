# Changelog

## 0.5.0

 * Update dependencies.
 * Fix build on recent node versions.
 * Drop support for node < 18.

## 0.4.5

 * Download and extract dmg image on Mac.

## 0.4.4

 * Download libxl via HTTPS
 * Upstream bug fix -> reenable ranges suite

## 0.4.3

 * Support Node 13
 * Temporarily disable tests relating to named ranges (upstream bug)

## 0.4.2

 * No changes, bogus release to keep 0.4.x the default on NPM.

## 0.4.1

 * Pin tar to 4.4.10 to avoid build break with 4.4.11

## 0.4.0

 * Fix build on node 12
 * Drop support for node < 10
 * Dependency bump

## 0.3.5

 * Download and extract dmg image on Mac.

## 0.3.4

 * Download libxl via HTTPS

## 0.3.3

 * Pin tar to 4.4.10 to avoid build break with 4.4.11

## 0.3.2

 * Fix build on node version 8.0 -- 8.9

## 0.3.1

 * Fix build on node 10
 * Fix deprecation warnings
 * Properly support async context and hooks

## 0.3.0

 * Build fixes for node 9
 * Remove support for node < 4 + io.js
 * Adapt to tar API changes

## 0.2.22

 * Download and extract dmg image on Mac.

## 0.2.21

 * Download libxl via HTTPS.

## 0.2.20

 * Download libxl via HTTP instead of FTP.

## 0.2.19

 * Allow formats belonging to a different book as templates for addFormat (see
    [issue #10](https://github.com/DirtyHairy/node-libxl/issues/10)).
 * Bump dependency versions.

## 0.2.18

 * Bump nan, fixes build on node 6.0.
 * Bump npm dependency versions.

## 0.2.17

 * Bump npm dependency versions.
 * Migrate from MD5 to md5.

## 0.2.16

 * Bump nan, fix resulting build issues -> io.js 3.2 support

## 0.2.15

 * Bump nan, fixes build on io.js 2.0.2
