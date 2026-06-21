# sales data cleaner and mis report generator

from pathlib import Path
import matplotlib.pyplot as plt
import pandas as pd


def sales_mis_report(file_path):
    try:
        print("\nLoading dataset...")

        extension = Path(file_path).suffix.lower()

        if extension == ".csv":
            df = pd.read_csv(file_path)

        elif extension in [".xlsx", ".xls"]:
            df = pd.read_excel(file_path)

        else:
            raise ValueError(
                f"Unsupported file format: {extension}"
            )

        print(f"Original Shape: {df.shape}")

        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.lower()
            .str.replace(" ", "_", regex=False)
            .str.replace("-", "_", regex=False)
        )

        print("\nDetected Columns:")
        print(df.columns.tolist())

        duplicates_removed = df.duplicated().sum()
        df = df.drop_duplicates()

        print(f"\nDuplicates Removed: {duplicates_removed}")

        numeric_cols = df.select_dtypes(include="number").columns

        for col in numeric_cols:
            df[col] = df[col].fillna(df[col].median())

        object_cols = df.select_dtypes(include=["object"]).columns

        for col in object_cols:
            df[col] = df[col].fillna("Unknown")

        possible_sales_cols = [
            "sales",
            "sales_amount",
            "revenue",
            "total_sales",
            "amount",
            "order_amount",
            "selling_price"
        ]

        sales_col = None

        for col in possible_sales_cols:
            if col in df.columns:
                sales_col = col
                break

        if sales_col is None:

            numeric_columns = (
                df.select_dtypes(include="number")
                .columns
                .tolist()
            )

            if numeric_columns:
                sales_col = numeric_columns[-1]

                print(
                    f"\nSales column not found."
                    f"\nAutomatically using: {sales_col}"
                )
            else:
                raise ValueError(
                    "No numeric column available for sales calculation."
                )

        print(f"Sales Column: {sales_col}")

        possible_product_cols = [
            "product",
            "product_name",
            "item",
            "item_name",
            "sku",
            "category"
        ]

        product_col = None

        for col in possible_product_cols:
            if col in df.columns:
                product_col = col
                break

        print(f"Product Column: {product_col}")

        possible_date_cols = [
            "date",
            "order_date",
            "sales_date",
            "purchase_date"
        ]

        date_col = None

        for col in possible_date_cols:
            if col in df.columns:
                date_col = col
                break

        print(f"Date Column: {date_col}")

        total_sales = df[sales_col].sum()

        print("\n========================")
        print("MIS REPORT")
        print("========================")
        print(f"Total Sales: ₹{total_sales:,.2f}")

        top_products = None

        if product_col:

            top_products = (
                df.groupby(product_col)[sales_col]
                .sum()
                .sort_values(ascending=False)
                .head(5)
            )

            print("\nTop 5 Products:")
            print(top_products)

        monthly_sales = None

        if date_col:

            df[date_col] = pd.to_datetime(
                df[date_col],
                errors="coerce"
            )

            monthly_sales = (
                df.groupby(
                    df[date_col].dt.to_period("M")
                )[sales_col]
                .sum()
            )

            print("\nMonthly Sales Summary:")
            print(monthly_sales)

        output_file = "cleaned_sales_report.xlsx"

        with pd.ExcelWriter(output_file,
                            engine="openpyxl") as writer:

            # Cleaned Data
            df.to_excel(
                writer,
                sheet_name="Cleaned_Data",
                index=False
            )

            dashboard = pd.DataFrame({
                "Metric": [
                    "Total Rows",
                    "Duplicates Removed",
                    "Total Sales"
                ],
                "Value": [
                    len(df),
                    duplicates_removed,
                    total_sales
                ]
            })

            dashboard.to_excel(
                writer,
                sheet_name="Dashboard",
                index=False
            )

            if top_products is not None:

                top_products.to_frame(
                    name="Sales"
                ).to_excel(
                    writer,
                    sheet_name="Top_Products"
                )

            if monthly_sales is not None:

                monthly_sales.to_frame(
                    name="Sales"
                ).to_excel(
                    writer,
                    sheet_name="Monthly_Sales"
                )

        print(f"\nExcel Report Saved: {output_file}")

        if top_products is not None:

            plt.figure(figsize=(10, 5))

            top_products.plot(
                kind="bar"
            )

            plt.title(
                "Top 5 Products By Sales"
            )

            plt.xlabel(
                "Product"
            )

            plt.ylabel(
                "Sales"
            )

            plt.tight_layout()

            plt.savefig(
                "top_products_bar_chart.png"
            )

            plt.close()

        if monthly_sales is not None:

            plt.figure(figsize=(10, 5))

            monthly_sales.plot(
                marker="o"
            )

            plt.title(
                "Monthly Sales Trend"
            )

            plt.xlabel(
                "Month"
            )

            plt.ylabel(
                "Sales"
            )

            plt.tight_layout()

            plt.savefig(
                "monthly_sales_line_chart.png"
            )

            plt.close()

        print("\nCharts Saved Successfully!")
        print("top_products_bar_chart.png")
        print("monthly_sales_line_chart.png")

        print("\nProject Completed Successfully!")

    except Exception as e:
        print(f"\nError: {e}")


if __name__ == "__main__":

    file_path = input(
        "Enter CSV/Excel File Path: "
    )

    sales_mis_report(file_path)
