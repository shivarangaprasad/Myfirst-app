"""Clean box-office data and export movie rankings to Excel.

Install dependencies with:
	python -m pip install pandas openpyxl
"""

import argparse
import re
from pathlib import Path

import pandas as pd


def normalize_column_name(column: object) -> str:
	"""Convert a source column name to lowercase snake_case."""
	return re.sub(r"[^a-z0-9]+", "_", str(column).strip().lower()).strip("_")


def read_source(path: Path) -> pd.DataFrame:
	"""Read an Excel workbook or CSV file based on its extension."""
	suffix = path.suffix.lower()
	if suffix in {".xlsx", ".xls"}:
		return pd.read_excel(path)
	if suffix == ".csv":
		return pd.read_csv(path)
	raise ValueError(f"Unsupported input file type: {path.suffix}")


def clean_and_rank(data: pd.DataFrame) -> pd.DataFrame:
	"""Clean source values and add independent dense ranks for votes and rating."""
	data = data.copy()
	data.columns = [normalize_column_name(column) for column in data.columns]

	if len(set(data.columns)) != len(data.columns):
		raise ValueError("Column names are not unique after normalization.")

	required_columns = {"movie", "votes", "imdb_rating"}
	missing_columns = required_columns.difference(data.columns)
	if missing_columns:
		missing = ", ".join(sorted(missing_columns))
		raise ValueError(f"Input is missing required columns: {missing}")

	for column in data.select_dtypes(include=["object", "string"]).columns:
		data[column] = data[column].astype("string").str.strip()
		data[column] = data[column].replace("", pd.NA)

	data = data.dropna(subset=["movie"])
	data["votes"] = pd.to_numeric(
		data["votes"].astype("string").str.replace(",", "", regex=False),
		errors="coerce",
	)
	data["imdb_rating"] = pd.to_numeric(data["imdb_rating"], errors="coerce")
	data = data.drop_duplicates()

	data["votes_rank"] = data["votes"].rank(
		method="dense", ascending=False
	).astype("Int64")
	data["imdb_rating_rank"] = data["imdb_rating"].rank(
		method="dense", ascending=False
	).astype("Int64")
	return data.sort_values(
		["votes", "imdb_rating", "movie"],
		ascending=[False, False, True],
		na_position="last",
	)


def main() -> None:
	parser = argparse.ArgumentParser(description=__doc__)
	parser.add_argument(
		"--input",
		type=Path,
		help="Input .xlsx, .xls, or .csv file (defaults to the project workbook).",
	)
	parser.add_argument(
		"--output",
		type=Path,
		default=Path(__file__).with_name("BoxOfficeCollections_ranked.xlsx"),
		help="Output .xlsx workbook (default: BoxOfficeCollections_ranked.xlsx).",
	)
	args = parser.parse_args()

	source = args.input or Path(__file__).with_name("BoxOfficeCollections.xlsx")
	if args.input is None and not source.exists():
		csv_fallback = source.with_suffix(".csv")
		if csv_fallback.exists():
			source = csv_fallback
			print(f"Workbook not found; using available CSV: {source}")
	if not source.is_file():
		raise FileNotFoundError(f"Input file not found: {source}")
	if args.output.suffix.lower() != ".xlsx":
		raise ValueError("Output file must have an .xlsx extension.")
	if args.output.resolve() == source.resolve():
		raise ValueError("Output file must be different from the input file.")

	ranked_data = clean_and_rank(read_source(source))
	votes_ranking = ranked_data.sort_values(
		["votes_rank", "movie"], ascending=[True, True], na_position="last"
	)
	rating_ranking = ranked_data.sort_values(
		["imdb_rating_rank", "movie"], ascending=[True, True], na_position="last"
	)
	with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
		votes_ranking.to_excel(writer, sheet_name="Votes Ranking", index=False)
		rating_ranking.to_excel(writer, sheet_name="IMDb Rating Ranking", index=False)

	print(f"Exported {len(ranked_data)} movies to {args.output}.")


if __name__ == "__main__":
	main()
