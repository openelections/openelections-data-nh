#!/usr/bin/env bash
# Run the OpenElections data-validation tests against the CSVs in this repo,
# the same way CI does. The test code lives in a separate repo
# (openelections/openelections-data-tests), pinned to a version in the CI
# workflow. We clone that repo into .data_tests/ on first run and re-use it.
#
# Usage:
#   scripts/run-data-tests.sh                # run all four test groups
#   scripts/run-data-tests.sh duplicate_entries
#   scripts/run-data-tests.sh missing_values 2018/20181106__nh__general__precinct.csv
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
WORKFLOW="$REPO_ROOT/.github/workflows/data_tests.yml"
CACHE_DIR="$REPO_ROOT/.data_tests"
ALL_TESTS=(file_format duplicate_entries missing_values vote_breakdown_totals)

# Pull the ref from the workflow so the local run can't drift from CI.
REF="$(awk '/repository:[[:space:]]*openelections\/openelections-data-tests/{found=1;next} found && /ref:/{print $2; exit}' "$WORKFLOW")"
if [[ -z "$REF" ]]; then
    echo "Could not read data-tests ref from $WORKFLOW" >&2
    exit 1
fi

if [[ ! -d "$CACHE_DIR/.git" ]]; then
    git clone --depth 1 --branch "$REF" \
        https://github.com/openelections/openelections-data-tests.git \
        "$CACHE_DIR"
elif [[ "$(git -C "$CACHE_DIR" describe --tags --exact-match 2>/dev/null || true)" != "$REF" ]]; then
    git -C "$CACHE_DIR" fetch --depth 1 origin "refs/tags/$REF:refs/tags/$REF"
    git -C "$CACHE_DIR" checkout "$REF"
fi

if [[ $# -eq 0 ]]; then
    TESTS=("${ALL_TESTS[@]}")
    FILES_ARGS=()
else
    TESTS=("$1")
    shift
    FILES_ARGS=()
    if [[ $# -gt 0 ]]; then
        FILES_ARGS=(--files "$@")
    fi
fi

status=0
for test in "${TESTS[@]}"; do
    echo "===> $test"
    if ! python3 "$CACHE_DIR/run_tests.py" "$test" "$REPO_ROOT" "${FILES_ARGS[@]}"; then
        status=1
    fi
done
exit $status
