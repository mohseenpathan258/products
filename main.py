import time

import flet as ft
import openpyxl


def main(page: ft.Page):
    # add records form
    def add_record_form(e):
        dialog_box.title = ft.Text(
            value="Add Records",
            weight=ft.FontWeight.BOLD
        )
        dialog_box.content = ft.Column(
            controls=[
                # Sr No label and textfield
                ft.Text(" Sr No"),
                ft.CupertinoTextField(
                    value=str(sh.max_row),
                    color=ft.Colors.GREY_400,
                    read_only=True,
                ),

                # space between Sr No and Goods Description
                ft.Container(height=5),

                # Goods Description label and textfield
                ft.Text(" Goods Description"),
                ft.CupertinoTextField(
                    placeholder_text="Goods Description",
                    placeholder_style=ft.TextStyle(
                        color=ft.Colors.GREY_400,
                    )
                ),

                # space between Goods Description and Part Number
                ft.Container(height=5),

                # Part Number label and textfield
                ft.Text(" Part Number"),
                ft.CupertinoTextField(
                    placeholder_text="Part Number",
                    placeholder_style=ft.TextStyle(
                        color=ft.Colors.GREY_400,
                    )
                ),

                # space between Part Number and Purchase Rate
                ft.Container(height=5),

                # Purchase Rate label and textfield
                ft.Text(" Purchase Rate"),
                ft.CupertinoTextField(
                    placeholder_text="Purchase Rate",
                    placeholder_style=ft.TextStyle(
                        color=ft.Colors.GREY_400,
                    ),
                    keyboard_type=ft.KeyboardType.NUMBER
                ),

                # space between Purchase Rate and Stock
                ft.Container(height=5),

                # Stock label and textfield
                ft.Text(" Stock"),
                ft.CupertinoTextField(
                    placeholder_text="Stock",
                    placeholder_style=ft.TextStyle(
                        color=ft.Colors.GREY_400,
                    ),
                    keyboard_type=ft.KeyboardType.NUMBER
                ),

                # space between Stock and HSN Code
                ft.Container(height=5),

                # HSN Code label and textfield
                ft.Text(" HSN Code"),
                ft.CupertinoTextField(
                    placeholder_text="HSN Code",
                    placeholder_style=ft.TextStyle(
                        color=ft.Colors.GREY_400,
                    ),
                    keyboard_type=ft.KeyboardType.NUMBER
                ),

                # space between HSN Code and GST
                ft.Container(height=5),

                # GST label and textfield
                ft.Text(" GST"),
                ft.CupertinoTextField(
                    placeholder_text="GST",
                    placeholder_style=ft.TextStyle(
                        color=ft.Colors.GREY_400,
                    ),
                    keyboard_type=ft.KeyboardType.NUMBER
                ),
            ]
        )

        # add button to add record to products_db.xlsx file
        dialog_box.actions.insert(
            0,
            ft.TextButton(
                text="Add",
                on_click=add_record_to_database
            )
        )

        page.open(dialog_box)

    # add record button
    def add_record_to_database(e):
        # add record to particular row (means after last row)
        select_row = sh.max_row + 1
        # Sr No
        sh.cell(row=select_row, column=1).value = int(dialog_box.content.controls[1].value)

        # Goods Description
        sh.cell(row=select_row, column=2).value = str(dialog_box.content.controls[4].value)

        # Part Number
        if str(dialog_box.content.controls[7].value).isdigit():
            sh.cell(row=select_row, column=3).value = int(dialog_box.content.controls[7].value)
        else:
            sh.cell(row=select_row, column=3).value = str(dialog_box.content.controls[7].value)

        # Purchase Rate
        sh.cell(row=select_row, column=4).value = int(dialog_box.content.controls[10].value)

        # Stock
        sh.cell(row=select_row, column=5).value = int(dialog_box.content.controls[13].value)

        # HSN Code
        sh.cell(row=select_row, column=6).value = int(dialog_box.content.controls[16].value)

        # GST
        sh.cell(row=select_row, column=7).value = int(dialog_box.content.controls[19].value)

        # save products_db
        wb.save("products_db.xlsx")

        # insert done icon at o index in dialog_box actions
        dialog_box.actions.insert(
            0,
            ft.Icon(
                name=ft.Icons.DONE,
                color=ft.Colors.GREEN
            )
        )

        # to update the new data record on the page
        add_data_row()

        page.update()

        # wait for one second before closing dialog box
        time.sleep(1)

        page.close(dialog_box)

    # page settings
    page.padding = 0
    page.spacing = 0
    page.bgcolor = ft.Colors.GREY_100

    # alert dialog box
    dialog_box = ft.AlertDialog(
        # dialog box background color
        surface_tint_color=ft.Colors.WHITE,
        scrollable=True,
        modal=True,
        actions=[
            ft.TextButton(
                text="Close",
                on_click=lambda e: page.close(dialog_box)
            )
        ]
    )

    # page appbar
    page.appbar = ft.AppBar(
        adaptive=True,
        # row for home icon and products title
        title=ft.Row(
            spacing=10, run_spacing=0,
            controls=[
                # home icon
                ft.Icon(
                    name=ft.Icons.HOME,
                    color=ft.Colors.WHITE,
                    size=30
                ),

                # products title
                ft.Text(
                    value="PRODUCTS",
                    color=ft.Colors.WHITE,
                    weight=ft.FontWeight.BOLD
                )
            ]
        ),

        # app background color
        bgcolor=ft.Colors.GREEN,

        # action button on the end of appbar
        actions=[
            # menu button
            ft.PopupMenuButton(
                icon=ft.Icons.MENU_OPEN,
                icon_color=ft.Colors.WHITE,
                items=[
                    # menu heading
                    ft.PopupMenuItem(
                        content=ft.Text(
                            value="Menu",
                            weight=ft.FontWeight.BOLD
                        )

                    ),

                    # divider
                    ft.PopupMenuItem(),

                    # add records to products table button
                    ft.PopupMenuItem(
                        content=ft.Row(
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            controls=[
                                # add records title
                                ft.Text(
                                    value="Add New Record"
                                ),

                                # icon button to show add record form
                                ft.IconButton(
                                    icon=ft.Icons.ADD_BOX,
                                    icon_color=ft.Colors.GREEN,
                                    on_click=add_record_form
                                ),
                            ]
                        )
                    ),

                    # divider
                    ft.PopupMenuItem(),

                    # control visibility of purchase rate column
                    ft.PopupMenuItem(
                        content=ft.Row(
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            controls=[
                                # visible purchase rate title
                                ft.Text(
                                    value="Visible Purchase Rate"
                                ),

                                # icon button to control visibility of purchase rate column
                                ft.IconButton(
                                    icon=ft.Icons.VISIBILITY,
                                    icon_color=ft.Colors.GREEN,
                                ),
                            ]
                        )

                    ),

                    # divider
                    ft.PopupMenuItem(),

                    # Information
                    ft.PopupMenuItem(
                        content=ft.Column(
                            spacing=20, run_spacing=0,
                            controls=[
                                # single click instruction
                                ft.Text(
                                    value="Single click to view row info",
                                    color=ft.Colors.GREY,
                                ),

                                # double click instruction
                                ft.Text(
                                    value="Double click to edit row info",
                                    color=ft.Colors.GREY,
                                )
                            ]
                        )
                    ),
                ]
            )
        ]
    )

    # search box
    search_box = ft.CupertinoTextField(
        placeholder_text="Search",
        border_radius=0,
        padding=ft.padding.only(
            left=20, top=15, bottom=15
        ),
        border=ft.border.only(
            bottom=ft.BorderSide(width=1, color=ft.Colors.GREY_300)
        ),
        suffix=ft.Container(
            ft.Icon(
                name=ft.Icons.CLOSE,
                color=ft.Colors.RED
            ),
            bgcolor=ft.Colors.WHITE,
            padding=ft.padding.only(right=20),
            on_click=None
        )
    )

    # data row
    data_row = ft.Column(
        spacing=1, run_spacing=0,
        scroll=ft.ScrollMode.HIDDEN,
        height=690
    )

    # products table container
    products_table = ft.Row(
        spacing=0, run_spacing=0,
        scroll=ft.ScrollMode.HIDDEN,
        controls=[
            ft.Column(
                spacing=1, run_spacing=0,
                controls=[
                    ft.Row(
                        spacing=1,
                        controls=[
                            # quote column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=100,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Quote",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # Sr No column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=100,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Sr No",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # goods description column
                            ft.Container(
                                padding=ft.padding.only(left=10, top=10, bottom=10),
                                width=300,
                                alignment=ft.alignment.center_left,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Goods Description",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # part number column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=200,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Part Number",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # purchase rate column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=150,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Purchase rate",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # sale rate column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=150,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Sale rate",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # stock column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=100,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Stock",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # hsn code column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=200,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="HSN Code",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # gst column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=100,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="GST",
                                    weight=ft.FontWeight.BOLD
                                )
                            ),

                            # total column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=150,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_300,
                                content=ft.Text(
                                    value="Total",
                                    weight=ft.FontWeight.BOLD
                                )
                            )
                        ]
                    ),

                    data_row
                ]
            )
        ]
    )

    # load workbook
    wb = openpyxl.load_workbook("products_db.xlsx")
    # select sheet
    sh = wb["Sheet1"]

    def add_data_row():
        # clear data_row before append
        data_row.controls.clear()

        for i in range(2, sh.max_row + 1):
            # sale rate calculation from purchase rate
            sale_rate = round(int(sh.cell(row=i, column=4).value) * 0.1 + int(sh.cell(row=i, column=4).value))

            # append data to data_row
            data_row.controls.append(
                ft.Row(
                    expand=True,
                    spacing=1, run_spacing=0,
                    controls=[
                        # quote column
                        ft.Container(
                            width=100,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Container(
                                width=25,
                                height=25,
                                border_radius=5,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREEN,
                                content=ft.Icon(
                                    name=ft.Icons.ADD,
                                    color=ft.Colors.WHITE
                                )
                            )
                        ),

                        # Sr No column
                        ft.Container(
                            width=100,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=str(sh.cell(row=i, column=1).value)
                            )
                        ),

                        # goods description column
                        ft.Container(
                            width=300,
                            height=60,
                            padding=ft.padding.only(left=10, right=10),
                            alignment=ft.alignment.center_left,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=str(sh.cell(row=i, column=2).value)
                            )
                        ),

                        # part number column
                        ft.Container(
                            width=200,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=str(sh.cell(row=i, column=3).value)
                            )
                        ),

                        # purchase rate column
                        ft.Container(
                            width=150,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=f"{sh.cell(row=i, column=4).value} \u20B9"
                            )
                        ),

                        # sale rate column
                        ft.Container(
                            width=150,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=f"{sale_rate} \u20B9"
                            )
                        ),

                        # stock column
                        ft.Container(
                            width=100,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=str(sh.cell(row=i, column=5).value)
                            )
                        ),

                        # hsn code column
                        ft.Container(
                            width=200,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=str(sh.cell(row=i, column=6).value)
                            )
                        ),

                        # gst column
                        ft.Container(
                            width=100,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=f"{sh.cell(row=i, column=7).value}%"
                            )
                        ),

                        # total column
                        ft.Container(
                            width=150,
                            height=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.WHITE,
                            content=ft.Text(
                                value=f"{round(sale_rate * 0.18 + sale_rate)} \u20B9"
                            )
                        )
                    ]
                )
            )

            # if stock is zero than show it in red color
            if sh.cell(row=i, column=5).value == 0:
                data_row.controls[i-2].controls[6].content.color = ft.Colors.RED

            # if no part number available show empty cell
            if sh.cell(row=i, column=3).value is None:
                data_row.controls[i - 2].controls[3].content.value = ""

    add_data_row()

    page.add(
        search_box,
        products_table
    )


ft.app(main)
