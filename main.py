import time
import flet as ft
from home import HomePage
from quotation import QuotationPage


def main(page: ft.Page):
    # to switch the pages
    def switch_page(page_name):
        # clear controls from the page
        # it won't clear the appbar
        page.controls.clear()
        # if page_name is "home" than it load home page {HomePage()}
        if page_name == "home":
            page.controls.append(HomePage(page, visibility_purchase_btn))
            # if page_name is "about" than it loads about page {AboutPage()}
        elif page_name == "quotation":
            page.controls.append(QuotationPage(page, visibility_purchase_btn))

        page.update()

    # it controls the visibility of purchase rate column
    def control_visibility_of_purchase_rate(e):
        # HomePage class instance
        home = HomePage(page, visibility_purchase_btn)
        home.add_product_content()

        # initially purchase rate is hidden
        # if we click on eye btn in appbar than it will be visible
        if home.visibility_purchase_btn.content.controls[0].name == ft.Icons.VISIBILITY_OFF:
            # change icon to visibility and color to green
            home.visibility_purchase_btn.content.controls[0].name = ft.Icons.VISIBILITY
            home.visibility_purchase_btn.content.controls[0].color = ft.Colors.GREEN

            for i in range(0, home.sh.max_row * 2 - 2):
                # heading of Purchase Rate visibility False
                home.products.controls[0].controls[0].controls[4].visible = True

                # content of Purchase Rate visibility False
                if i % 2 == 0:
                    home.products_content.controls[i].controls[4].visible = True
                else:
                    # increase the width of divider of products content
                    home.products_content.controls[i].width = 1378
        else:
            # change icon to visibility_off and color to red
            home.visibility_purchase_btn.content.controls[0].name = ft.Icons.VISIBILITY_OFF
            home.visibility_purchase_btn.content.controls[0].color = ft.Colors.RED

            for i in range(0, home.sh.max_row * 2 - 2):
                # heading of Purchase Rate visibility True
                home.products.controls[0].controls[0].controls[4].visible = False

                # content of Purchase Rate visibility True
                if i % 2 == 0:
                    home.products_content.controls[i].controls[4].visible = False
                else:
                    # decrease the width of divider of products content
                    home.products_content.controls[i].width = 1228
        page.update()

    # clear the quotation
    def clear_quotation(e):
        # QuotationPage class instance
        quotation = QuotationPage(page, visibility_purchase_btn)
        quotation.quotation_content.controls.clear()  # quotation content cleared

        # HomePage class instance
        home = HomePage(page, visibility_purchase_btn)
        home.quotation_list.clear()  # quotation list cleared

        snack_bar.content.value = "Quotation cleared"
        page.open(snack_bar)

        page.update()

    # add record button
    def add_record_to_database(e):
        # HomePage class instance
        home = HomePage(page, visibility_purchase_btn)

        # add record to particular row (means after last row) 
        select_row = home.sh.max_row + 1

        # Sr No
        home.sh.cell(row=select_row, column=1).value = int(dialog_box.content.controls[1].value)

        # Goods Description
        home.sh.cell(row=select_row, column=2).value = str(dialog_box.content.controls[4].value)

        # Part Number
        if str(dialog_box.content.controls[7].value).isdigit():
            home.sh.cell(row=select_row, column=3).value = int(dialog_box.content.controls[7].value)
        else:
            home.sh.cell(row=select_row, column=3).value = str(dialog_box.content.controls[7].value)

        # Purchase Rate
        home.sh.cell(row=select_row, column=4).value = int(dialog_box.content.controls[10].value)

        # Stock
        home.sh.cell(row=select_row, column=5).value = int(dialog_box.content.controls[13].value)

        # HSN Code
        home.sh.cell(row=select_row, column=6).value = int(dialog_box.content.controls[16].value)

        # GST
        home.sh.cell(row=select_row, column=7).value = int(dialog_box.content.controls[19].value)

        # save products_db
        home.wb.save("products_db.xlsx")

        # insert done icon at o index in dialog_box actions
        dialog_box.actions.insert(
            0,
            ft.Icon(
                name=ft.Icons.DONE,
                color=ft.Colors.GREEN
            )
        )

        # to update the new data record on the page
        home.add_product_content()

        page.update()

        # wait for one second before closing dialog box
        time.sleep(1)

        # close dialog_box
        page.close(dialog_box)
        # clear all button from dialog_box
        dialog_box.actions.clear()

    # add records form
    def add_record_form(e):
        # HomePage class instance
        home = HomePage(page, visibility_purchase_btn)

        dialog_box.title = ft.Text(
            value="Add Records",
            weight=ft.FontWeight.BOLD
        )
        dialog_box.content = ft.Column(
            controls=[
                # Sr No label and textfield
                ft.Text(" Sr No"),
                ft.CupertinoTextField(
                    value=str(home.sh.max_row),
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
        # first clear if add buttons are already added
        dialog_box.actions.clear()
        dialog_box.actions.append(
            ft.TextButton(
                text="Add",
                on_click=add_record_to_database
            )
        )
        dialog_box.actions.append(
            ft.TextButton(
                text="Close",
                on_click=lambda event: page.close(dialog_box)
            )
        )

        page.open(dialog_box)

    # page settings
    page.spacing = 0
    page.padding = 0
    page.bgcolor = ft.Colors.WHITE

    # alert dialog box
    dialog_box = ft.AlertDialog(
        # dialog box background color
        surface_tint_color=ft.Colors.WHITE,
        scrollable=True,
        modal=True,
    )

    # snack bar
    snack_bar = ft.SnackBar(
        content=ft.Text()
    )

    # purchase rate visibility btn
    visibility_purchase_btn = ft.PopupMenuItem(
        content=ft.Row(
            controls=[
                # visible purchase rate icon
                ft.Icon(
                    name=ft.Icons.VISIBILITY_OFF,
                    color=ft.Colors.RED
                ),

                # visible purchase rate title
                ft.Text(
                    value="Visible Purchase Rate",
                    color=ft.Colors.BLACK,
                )
            ]
        ),
        on_click=control_visibility_of_purchase_rate
    )

    # page appbar
    page.appbar = ft.AppBar(
        bgcolor="#008FD5",
        toolbar_height=50,

        # appbar HR logo and Rajasthan Hydraulics title
        title=ft.Row(
            controls=[
                ft.Container(
                    width=35,
                    height=35,
                    # padding=5,
                    bgcolor=ft.Colors.WHITE,
                    border_radius=20,
                    # content=ft.Image(
                    #     src="logo.png"
                    # ),
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=0, run_spacing=0,
                        controls=[
                            ft.Text(value="H", scale=0.8, color=ft.Colors.BLACK, weight=ft.FontWeight.BOLD),
                            ft.Text(value="R", scale=0.8, color="#008FD5", weight=ft.FontWeight.BOLD)
                        ]
                    ),
                    on_click=lambda e: switch_page("home")
                ),

                ft.Text(
                    value="RAJASTHAN HYDRAULICS",
                    color=ft.Colors.WHITE,
                    size=18,
                    weight=ft.FontWeight.W_500
                )
            ]
        ),

        actions=[
            ft.PopupMenuButton(
                bgcolor=ft.Colors.WHITE,
                # this does not work in flet app in mobile
                # content=ft.Image(
                #     src="menu_icon.svg",
                #     scale=0.6,
                #     color=ft.Colors.WHITE
                # ),

                icon=ft.Icons.MENU,
                icon_color=ft.Colors.WHITE, 

                items=[
                    # menu heading
                    ft.PopupMenuItem(
                        content=ft.Text(
                            value="Menu",
                            color=ft.Colors.BLACK,
                            weight=ft.FontWeight.W_600
                        ),
                        mouse_cursor=ft.MouseCursor.BASIC
                    ),

                    # divider
                    ft.PopupMenuItem(),

                    # go to quotation page btn
                    ft.PopupMenuItem(
                        content=ft.Row(
                            controls=[
                                # icon button to show quotation
                                ft.Icon(
                                    name=ft.Icons.REQUEST_QUOTE,
                                    color="#008FD5"
                                ),

                                # quotation title
                                ft.Text(
                                    value="View Quotation",
                                    color=ft.Colors.BLACK
                                ),
                            ]
                        ),
                        on_click=lambda e: switch_page("quotation")
                    ),

                    # clear quotation
                    ft.PopupMenuItem(
                        content=ft.Row(
                            controls=[
                                # clear quotation icon
                                ft.Icon(
                                    name=ft.Icons.REQUEST_QUOTE,
                                    color=ft.Colors.RED
                                ),

                                # clear quotation title
                                ft.Text(
                                    value="Clear Quotation",
                                    color=ft.Colors.BLACK
                                )
                            ]
                        ),
                        on_click=clear_quotation
                    ),

                    # add new record
                    ft.PopupMenuItem(
                        content=ft.Row(
                            controls=[
                                # add new record icon
                                ft.Icon(
                                    name=ft.Icons.ADD_BOX,
                                    color=ft.Colors.GREEN
                                ),

                                # add new record title
                                ft.Text(
                                    value="Add New Record",
                                    color=ft.Colors.BLACK,
                                )
                            ]
                        ),
                        on_click=add_record_form
                    ),

                    # visible purchase rate
                    visibility_purchase_btn,

                    # information of "on tap" and "on long press"
                    ft.PopupMenuItem(
                        content=ft.Column(
                            spacing=20, run_spacing=0,
                            controls=[
                                # information title
                                ft.Text(
                                    value="Information",
                                    color=ft.Colors.BLACK,
                                    weight=ft.FontWeight.W_600
                                ),

                                # single click instruction
                                ft.Text(
                                    value="Tap to view Record",
                                    color=ft.Colors.GREY,
                                ),

                                # double click instruction
                                ft.Text(
                                    value="Long press to edit Record",
                                    color=ft.Colors.GREY,
                                )
                            ]
                        )
                    )
                ]
            ),

            # to provide some space to menu button from right side in the appbar
            ft.Container(
                width=10,
                bgcolor="#008FD5"
            )
        ]
    )

    # initially load home page (HomePage())
    switch_page("home")


ft.app(target=main)
