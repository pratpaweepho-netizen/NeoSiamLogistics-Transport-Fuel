#!/bin/bash
# NeoSiam fleet-fuel: nightly copy of the WHOLE Firebase database to the Synology NAS (user decision 19 Nov 2026).
# Runs in DSM: Control Panel > Task Scheduler > Create > Scheduled Task > User-defined script (user: root, daily 02:00).
# Paste this whole file into "Run command". Only the 3 lines under SETTINGS may need changing.
#
# What it does: download the database (read only - it never writes to Firebase), check the copy is complete
# (size, starts with { ends with }, parses as JSON with python3 or jq, trips by month + settings), gzip it as
# fleet-fuel_YYYY-MM-DD_HHMM.json.gz, test the gzip, delete copies older than KEEP_DAYS, write one line to backup.log.
# Any problem -> exit 1, so DSM emails the run details ("only when the script terminates abnormally").
# Restore: download the .json.gz from File Station, then in the app: Import Excel page > the full backup / restore
# card > pick the file (the app unpacks .gz itself). The copies hold driver bank account numbers: give the shared
# folder to admins only.
#
# NOTE: this works because the database can be read without a login today. If the database rules are ever
# locked (security phase, parked by the user), add an access token to DB_URL/.json?auth=... or the run fails
# with "download failed" - the email will say so; nothing is deleted on a failed run.
# Plain ASCII on purpose: the DSM script box does not handle Thai.

# ---- SETTINGS ----
DEST="/volume1/backup/fleet-fuel"   # shared folder "backup" + sub folder; change if you use another shared folder
KEEP_DAYS=60                         # copies older than this are deleted (only after a successful run)
DB_URL="https://fleet-fuel-management-system-default-rtdb.asia-southeast1.firebasedatabase.app"
# ------------------

MIN_BYTES=1000000   # the database is ~16 MB; anything under 1 MB is a broken / empty download
set -u
umask 022

SHARE="$(echo "$DEST" | cut -d/ -f1-3)"   # /volume1/backup
[ -d "$SHARE" ] || { echo "ERROR: shared folder $SHARE does not exist - create it in Control Panel > Shared Folder"; exit 1; }
mkdir -p "$DEST" || { echo "ERROR: cannot create $DEST"; exit 1; }

STAMP="$(date +%Y-%m-%d_%H%M)"
TMP="$DEST/.incoming_$STAMP.json"
OUT="$DEST/fleet-fuel_$STAMP.json.gz"
LOG="$DEST/backup.log"

fail() {
    echo "$(date '+%Y-%m-%d %H:%M:%S') FAIL $1" >> "$LOG"
    echo "ERROR: $1"
    rm -f "$TMP" "$OUT.part"
    exit 1
}

curl -sS -f --retry 3 --retry-delay 30 --connect-timeout 30 --max-time 900 -o "$TMP" "$DB_URL/.json"
RC=$?
[ $RC -eq 0 ] || fail "download failed (curl exit $RC)"

SIZE=$(wc -c < "$TMP" | tr -d ' ')
[ "$SIZE" -ge "$MIN_BYTES" ] || fail "download too small ($SIZE bytes) - not a full database copy"
[ "$(head -c 1 "$TMP")" = "{" ] || fail "download does not start with { - not JSON"
[ "$(tail -c 1 "$TMP")" = "}" ] || fail "download does not end with } - cut off"
grep -q '"trips":' "$TMP" || fail "no trips part in the copy"
grep -q '"settings":' "$TMP" || fail "no settings part in the copy"

# Parse it as JSON and check the shape (root object, trips = months YYYY-MM with records, settings object).
# Prints the number of trips. Exit 2 = no JSON tool on this NAS (then old copies are NOT deleted, see below).
check_json() {
    if command -v python3 >/dev/null 2>&1; then
        python3 - "$TMP" <<'PY'
import json, re, sys
d = json.load(open(sys.argv[1], encoding='utf-8'))
if not isinstance(d, dict): sys.exit('root is not an object')
t = d.get('trips')
if not isinstance(t, dict) or not t: sys.exit('trips missing or empty')
if not all(re.match(r'^[0-9]{4}-[0-9]{2}$', k) and isinstance(v, (dict, list)) for k, v in t.items()): sys.exit('trips not by month')
if not isinstance(d.get('settings'), dict): sys.exit('settings missing')
n = sum(len([x for x in (m.values() if isinstance(m, dict) else m) if x is not None]) for m in t.values())
if n < 1: sys.exit('no trips')
print(n)
PY
    elif command -v jq >/dev/null 2>&1; then
        # same rules and the same count as the python branch (non-null records in every month)
        N="$(jq -e 'if type == "object" and (.trips | type == "object" and length > 0)
                      and (.trips | to_entries | all((.key | test("^[0-9]{4}-[0-9]{2}$"))
                                                     and (.value | type == "object" or type == "array")))
                      and (.settings | type == "object")
                    then [.trips[] | [.[] | select(. != null)] | length] | add
                    else error("shape") end' "$TMP")" || { echo "shape check failed" >&2; return 1; }
        [ "$N" -ge 1 ] 2>/dev/null || { echo "no trips" >&2; return 1; }
        echo "$N"
    else
        return 2
    fi
}
TRIPS="$(check_json)"
RC=$?
PARSED=yes
if [ $RC -eq 2 ]; then
    PARSED=no
elif [ $RC -ne 0 ]; then
    fail "the copy is not a valid database (JSON / shape check failed)"
fi

gzip -c "$TMP" > "$OUT.part" || fail "gzip failed (disk full?)"
gzip -t "$OUT.part" || fail "gzip test failed"
mv "$OUT.part" "$OUT" || fail "cannot rename $OUT.part"
rm -f "$TMP"

# old copies go only after today's copy is safely written AND was proven to be a valid database
if [ "$PARSED" = yes ]; then
    find "$DEST" -maxdepth 1 -name 'fleet-fuel_*.json.gz' -mtime +"$KEEP_DAYS" -delete
fi
find "$DEST" -maxdepth 1 -name '.incoming_*.json' -mtime +1 -delete
COUNT=$(ls -1 "$DEST"/fleet-fuel_*.json.gz 2>/dev/null | wc -l | tr -d ' ')
GZ=$(wc -c < "$OUT" | tr -d ' ')

if [ "$PARSED" = no ]; then
    # still a usable copy (text checks + gzip passed), but without a JSON tool nothing old is deleted
    echo "$(date '+%Y-%m-%d %H:%M:%S') WARN $(basename "$OUT") json=$SIZE gz=$GZ copies=$COUNT - no python3/jq: not parsed, old copies kept" >> "$LOG"
    echo "WARN $(basename "$OUT") kept, but this NAS has no python3 or jq: the copy was not parsed and old copies are not deleted"
    exit 1
fi
echo "$(date '+%Y-%m-%d %H:%M:%S') OK $(basename "$OUT") json=$SIZE gz=$GZ trips=$TRIPS copies=$COUNT" >> "$LOG"
echo "OK $(basename "$OUT") ($GZ bytes, $TRIPS trips, $COUNT copies kept)"
exit 0
