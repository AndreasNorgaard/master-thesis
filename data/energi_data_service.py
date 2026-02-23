import json
from pathlib import Path

import polars as pl
import requests


class EnergiDataServiceAPIClient:
    """
    Client for the Energi Data Service API.
    Source: https://www.energidataservice.dk/datasets

    Args:
        start_date: Start date of the data to fetch.
        end_date: End date of the data to fetch.
        price_area: Price area to fetch data for.

    Methods:
        day_ahead_prices: Fetches day ahead prices and returns them as a DataFrame.
        co2_emissions: Fetches CO2 emissions and returns them as a DataFrame.
    """

    def __init__(self, start_date: str, end_date: str, price_area: str | None = None):
        self.start_date = start_date
        self.end_date = end_date
        self.price_area = price_area

    def _get_response(self, url: str, output_path: Path | None) -> pl.DataFrame:
        """
        Make GET request to API and write response to file.

        Args:
            url: URL to fetch data from.
            output_path: Path to write response to.
        """
        print(f"Fetching data for period: '{self.start_date}' to '{self.end_date}'...")
        # Make GET request to API
        response = requests.get(
            url=url,
            params={
                "start": self.start_date,
                "end": self.end_date,
                "filter": json.dumps({"PriceArea": [self.price_area]})
                if self.price_area
                else None,
            },
        )

        # Write response to file
        records = response.json().get("records", [])
        df = pl.DataFrame(records)

        # Write DataFrame to Excel file
        if output_path:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            df.write_excel(output_path)

        return df

    def day_ahead_prices(self, write_to_file: bool = True) -> pl.DataFrame:
        """
        Fetches day ahead prices and returns them as a DataFrame.

        Args:
            write_to_file: Whether to write the data to an Excel file.
        """
        output_path = Path("data/raw/day_ahead_prices.xlsx") if write_to_file else None
        url = "https://api.energidataservice.dk/dataset/DayAheadPrices"
        df = self._get_response(url, output_path)
        print("Day ahead prices have been fetched!")
        return df

    def co2_emissions(self, write_to_file: bool = True) -> pl.DataFrame:
        """
        Fetches CO2 emissions and returns them as a DataFrame.

        Args:
            write_to_file: Whether to write the data to an Excel file.
        """
        output_path = Path("data/raw/co2_emissions.xlsx") if write_to_file else None
        url = "https://api.energidataservice.dk/dataset/CO2Emis"
        df = self._get_response(url, output_path)
        print("CO2 emissions have been fetched!")
        return df


if __name__ == "__main__":
    client = EnergiDataServiceAPIClient(
        start_date="2026-01-01",
        end_date="2026-01-31",
        price_area=None,
    )
    client.day_ahead_prices()
    client.co2_emissions()
