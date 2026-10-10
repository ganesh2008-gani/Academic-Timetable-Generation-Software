[README.md](https://github.com/user-attachments/files/33278250/README.md)
# Timetable Generator

An Excel/VBA tool that builds weekly class timetables for multiple programs (e.g. `BTCE-4`, `MTSE-2`) from a course list and each faculty member's availability. Placement is randomised and rule-based: theory hours are placed one period at a time, lab hours as two consecutive periods, and a faculty member is never booked in two places at once.

## Repository contents

| File | Description |
|---|---|
| `Timetable_Generator.xlsm` | Macro-enabled workbook. All logic lives in `Module1` (VBA). |
| `Sample_Input.xlsx` | Example input: time slots, faculty availability, course–faculty mapping. |
| `Sample_Output.xlsx` | Example output: one timetable sheet per program-semester, plus a summary sheet. |

## Weekly grid

6 days (MON–SAT). Mon–Fri have 7 teaching periods; Saturday has 4 (morning only).

| Period | 1 | 2 | | 3 | 4 | | 5 | 6 | 7 |
|---|---|---|---|---|---|---|---|---|---|
| Time | 8–9 | 9–10 | BREAK | 10:30–11:30 | 11:30–12:30 | LUNCH | 2–3 | 3–4 | 4–5 |

## Input format

The macro reads two sheets. The names are fixed in the code: **`COURSE_DATA`** and **`FACULTY_AVAIL`**. In `Sample_Input.xlsx` the equivalent sheets are named `Course-Faculty` and `Faculty-Availability`, so copy their contents into sheets with the macro's names. The `Times` sheet is for reference only; the macro does not read it.

### `COURSE_DATA` (from `Course-Faculty`)

One row per schedulable item. Row 1 is the header. Column positions matter:

| Col | Field | Notes |
|---|---|---|
| A | Program-Sem | Must match the name of an output sheet, e.g. `BTCE-4`. Rows with a blank value are never scheduled. |
| B | Course | Course code. Lab batch rows use a `-LB1`, `-LB2`… suffix (or `-B1`, `-B2`…). |
| C | Course Title | Informational. |
| D | L | Lecture hours per week. |
| E | T | Tutorial hours per week. Theory hours = L + T. |
| F | P | Practical hours per week. If P > 0 the row is treated as a lab and placed as 2 consecutive periods at a time. |
| G | (unused) | |
| H | NoB | Number of lab batches; only used by `ReplicateOnlyLabs` (see below). |
| I | Faculty1 | Faculty code. Required. |
| J | Faculty2 | Optional. Used for labs only. |

Faculty3–5 (columns K–M in the sample) are not read.

### `FACULTY_AVAIL` (from `Faculty-Availability`)

Row 1 is the header. Column A is the faculty code. Columns B onward are availability flags, **any non-blank value means available**:

- Columns B–H: MON #1–#7, I–O: TUE, P–V: WED, W–AC: THU, AD–AJ: FRI
- Columns AK–AN: SAT #1–#4

A faculty code that is missing from this sheet is treated as never available.

## Output format

One sheet per program-semester with the layout above. Each cell holds a course code. Colours:

- Green: theory
- Blue: lab
- Light yellow: break
- Dark yellow: lunch

See `Sample_Output.xlsx`. Its `Summary` sheet lists hours required, placed and left unplaced per program.

## Scheduling rules

Implemented in `GenerateSheet`:

1. Every course hour must be placed; a day and period are picked at random until a valid slot is found (up to 50,000 attempts per hour; if none is found, a "Unable to place subject" message appears).
2. A slot is valid only if the cell is empty and the faculty is available (`FACULTY_AVAIL`) and not already booked in another program at that time.
3. Labs need two adjacent free periods; both faculty (if two are listed) must be free for both.
4. At most 3 periods of the same course code per day.
5. Saturday: theory only in periods 1–4, labs start no later than period 3.

## Helper macros

| Macro | Purpose |
|---|---|
| `ReplicateOnlyLabs` | Deletes rows in `COURSE_DATA` whose course code ends in `-LB<n>`, then re-creates `-LB1`…`-LB<NoB>` copies of each lab row with `NoB > 1`. |
| `SetupHeaders(ws)` | Clears a sheet and writes the day/period header grid. |
| `GenerateSheet(ws)` | Fills one program's timetable. |

## Known issues and limitations

These came up while reviewing the code against `Sample_Input.xlsx`:

1. **No entry-point macro.** Nothing creates the program sheets or calls `SetupHeaders` and `GenerateSheet` for each program, so the generator cannot be run end-to-end from the workbook as it stands.
2. **Faculty clash tracking is inactive.** `facultyNames` and `facultyCount` are never populated, so `GetFacultyIndex` always returns 0 and faculty are never marked busy. Until this is fixed, the same faculty member can be scheduled in two programs at the same time.
3. **`ReplicateOnlyLabs` conflicts with the sample input.** The sample already contains expanded `-LB1..-LB4` rows with no base lab row. Running the macro deletes those rows and does not recreate them. Use it only with input in the "one lab row + `NoB`" format, or skip it.
4. **Lab batches run one after another.** Each batch takes its own slot in the same timetable. If batches actually run in parallel, `BTCE-4` in the sample needs 53 hours against 39 available slots and cannot fit, so 14 hours are left unplaced.
5. **Labs can straddle the break or lunch.** Periods 2→3 and 4→5 are treated as adjacent.
6. **Randomness.** `Randomize` is never called, so each Excel session repeats the same sequence. `Sample_Output.xlsx` was generated with a fixed seed (42), so a real run will produce different layouts.
7. **Sample data:** the `S2CTML1` (Design Studio) row has a blank Program-Sem and is skipped. Faculty2 is ignored for theory courses.



