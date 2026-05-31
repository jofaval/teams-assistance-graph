# Changelog

All notable changes to this project will be documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## 2026-05-31

## Added

- Error handling and user feedback for unsupported file formats or parsing issues, with clear messages and guidance for users to provide correct input.

### Modified

- Reset zoom is now more coherent to the rest of the UI.
- New color scheme for the graphs, with a more modern and clean look, and better contrast for accessibility.

### Fixed

- Selecting a member and searching, unselects the member and shows all the attendees again. Now it keeps the selected member and filters the attendees by the search query.

## 2026-05-26

### Added

- Added `AGENTS.md` with contributor procedure and changelog guidance.
- New attendees list view in the dashboard (`index.html`): searchable, sortable table with duration presets, filter chips, and a restore/list toggle.
- New UI elements: `restoreListButton`, filters bar and clipboard chip icon.
- Utility functions: `debounce`, `computeTimeSeriesForAttendees`, `mapParticipantNameToIdentityKey`, `setHeaderHeightVar`, `fillAttendeesList`, and `prepareAttendeesListView`.

### Modified

- Translated Spanish UI strings and ARIA labels to English across `index.html` and `main.js`.
- Improved attendee filtering and preset toggles, added a clear-all filters action and sync between presets and the input value.
- Integrated attendees list with graph range selection; added responsive two-column layout handling and header height adjustments.
- Console/debug messages switched to English and several internal methods refactored for clarity and reusability.

## 2026-05-09

### Modified

- Redesigned the dashboard layout with a clearer executive summary, improved accessibility attributes, and responsive behavior for mobile.
- Extended report parsing to support section-based extraction (participants, activities, interactions) with more robust date and duration handling.
- Added timeline controls to analyze attendance by minute/second and by custom time range, keeping graph and KPIs synchronized.
- Introduced richer engagement insights, including reactions metrics, dual distribution views, and enhanced tooltip/trend/critical-drop visualization.
* Improved tooltip hover

## 2026-05-08

### Added

* Stay distribution chart with quartiles and hover details
* Critical drop visualization with a red line and hover details

### Modified

- New visual identity for the dashboard, with a more modern and clean design

## 2026-02-11

### Added

* Hover displays attendees as well as the retention percentage at that time

### Disregarded

* A threshold line for the retention percentage, it is not dynamic

## 2025-05-26

### Added

- You can now see a distribution chart by quartiles, and hover to see the total number of attendees at that quartile

## 2025-05-17

### Modified

- General stats and clipboard are now updated when zooming in and out 

## 2025-05-13

### Added

- You can now visualize total watch time in hours and minutes

## 2025-05-12

### Modified

- It now can use an english report file to generate the graphs

## 2025-05-11

### Added

- Microsoft Teams Favicon
- Data truth notice
- Copy to clipboard button for the attendees over X minutes
- You can now click on a point in time to get the attendees at that time
- License
- You can now see the general stats for that meeting (avg retention, total attendees, etc.)

### Modified

- Refactored the code to use classes
- Chart title is now the Meeting title
- KPIs use Microsoft Teams Brand color

## 2025-05-10

### Added

- Base dashboard
    - Area chart
- Graphs for attendance report
- Query attendees over X minutes