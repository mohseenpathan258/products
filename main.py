import flet as ft
import openpyxl


def main(page: ft.Page):
    wb = openpyxl.load_workbook("db.xlsx")
    sh = wb["Sheet1"]

    cell_a2 = sh["A2"]
    cell_b2 = sh["B2"]

    cell_a3 = sh["A3"]
    cell_b3 = sh["B3"]

    cell_a4 = sh["A4"]
    cell_b4 = sh["B4"]

    table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("Name")),
            ft.DataColumn(ft.Text("Roll Number"))
        ],

        rows=[
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(cell_a2.value)),
                    ft.DataCell(ft.Text(cell_b2.value))
                ]
            ),

            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(cell_a3.value)),
                    ft.DataCell(ft.Text(cell_b3.value))
                ]
            ),

            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(cell_a4.value)),
                    ft.DataCell(ft.Text(cell_b4.value))
                ]
            )
        ]
    )

    page.add(table)


ft.app(target=main, assets_dir="assets")
