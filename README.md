[![Build Status](https://github.com/openelections/openelections-data-nh/actions/workflows/data_tests.yml/badge.svg?branch=master)](https://github.com/openelections/openelections-data-nh/actions/workflows/data_tests.yml?query=branch%3Amaster)

# OpenElections Data New Hampshire

This repository contains pre-processed election results from New Hampshire, formatted to be ingested into the OpenElections [processing pipeline](http://docs.openelections.net/guide/). It contains mostly CSV files converted from PDF tables. Interested in contributing? We have a bunch of [easy tasks](https://github.com/openelections/openelections-data-nh/labels/easy%20task) for you to tackle.

Here is what a [finished CSV file (from Ohio)](https://github.com/openelections/openelections-data-oh/blob/master/2000/20001107__oh__general__president.csv) looks like. Note that each row represents a single result for a single candidate, even if the data has multiple candidates in a single row. Also, vote totals do not contain commas or other formatting.

For extracting text from PDF tables, we recommend [Tabula](http://tabula.technology/), which can be installed and run locally on OSX, Windows or Linux.

If you're familiar with git and Github, clone this repository and get started. If not, you can still help: leave a comment on a task you'd like to work on, or just convert any of the files into CSV and send the result to openelections@gmail.com.

## Conventions for the modern parser framework

The 2022+ output CSVs (produced by `oe_nh/`) follow these conventions:

- **Output schema** is `county,precinct,office,district,party,candidate,votes` — one row per (precinct, candidate) pair.
- **Office names** are: `President`, `Governor`, `US Senate`, `Congressional`, `Executive Council`, `State Senate`, `State Representative`.
- **Floterial House districts** are emitted with an `F` suffix on the district column — e.g. Belknap's two-seat floterial that overlays Districts 1–7 appears as `district="8F"`. The SoS source file marks these with either `F` or `FL` on the district header (`District No. 8 (2) F` or `District No. 14 (1) FL`); both are normalized to `F` in output.
- **Recount columns** that the SoS sometimes interleaves alongside certified counts (e.g. inline `Recount` columns in some 2022 House districts; whole `RECOUNT FIGURES` duplicate district sections in 2024 Strafford) are **dropped**. We ship certified counts only. To get recount data, go to the SoS source files.
- **`BLC` columns** that appear in a handful of 2022 Rockingham House districts alongside Recount columns (an auxiliary ballot-related count whose meaning we haven't confirmed) are also **dropped**. If you know what BLC stands for and want it surfaced, open an issue.
- **Write-Ins / Undervotes / Overvotes** (in 2024 SoS files) are emitted as candidate rows with empty `party`. The `candidate` value is the literal column header (`Write-Ins`, `Undervotes`, `Overvotes`).
- **Source-file naming and per-office shape details** live in [scripts/fetch-raw.md](scripts/fetch-raw.md).
