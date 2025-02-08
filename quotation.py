import time
import flet as ft


class Quotation(ft.Column):
    def __init__(self, page):
        super(Quotation, self).__init__()
        self.page = page

        # snack bar
        self.snack_bar = ft.SnackBar(
            content=ft.Text()
        )

        self.app_bar = ft.ResponsiveRow(
            spacing=0, run_spacing=0,
            controls=[
                ft.Container(
                    height=70,
                    bgcolor=ft.Colors.YELLOW,
                    alignment=ft.alignment.center_left,
                    content=ft.Row(
                        controls=[
                            # app_bar home button
                            ft.IconButton(
                                icon=ft.Icons.ARROW_BACK_IOS_NEW,
                                icon_color=ft.Colors.BLACK,
                                icon_size=25,
                                on_click=lambda e: self.page.go("/products")
                            ),

                            # app_bar products title
                            ft.Text(
                                value="QUOTATION",
                                color=ft.Colors.BLACK,
                                size=25,
                                weight=ft.FontWeight.BOLD
                            ),

                            # quote clear button
                            ft.Row(
                                expand=True,
                                alignment=ft.MainAxisAlignment.END,
                                controls=[
                                    ft.IconButton(
                                        icon=ft.Icons.REQUEST_QUOTE,
                                        icon_color=ft.Colors.RED,
                                        on_click=self.clear_quotation
                                    )
                                ]
                            )
                        ]
                    )
                )
            ]
        )

        # quotation table
        self.quotation_table = ft.ResponsiveRow(
            spacing=0, run_spacing=1,
            controls=[
                # space between app_bar and quotation row
                ft.Container(height=10),

                ft.Row(
                    expand=True,
                    spacing=0,
                    controls=[
                        # quotation title row
                        ft.Container(
                            expand=True,
                            padding=ft.padding.only(top=15, bottom=15),
                            # width=100,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.GREY_300,
                            content=ft.Text(
                                value="Quotation",
                                size=20,
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.BOLD
                            )
                        )
                    ]
                ),

                # data row titles
                ft.Row(
                    expand=True,
                    spacing=1,
                    controls=[
                        # Sr No column
                        ft.Container(
                            padding=ft.padding.only(top=15, bottom=15),
                            width=60,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.GREY_200,
                            content=ft.Text(
                                value="Sr No",
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_600
                            )
                        ),

                        # Goods Description column
                        ft.Container(
                            padding=ft.padding.only(top=15, bottom=15, left=10, right=10),
                            width=200,
                            alignment=ft.alignment.center_left,
                            bgcolor=ft.Colors.GREY_200,
                            content=ft.Text(
                                value="Goods Description",
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_600
                            )
                        ),

                        # Quantity column
                        ft.Container(
                            padding=ft.padding.only(top=15, bottom=15),
                            width=80,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.GREY_200,
                            content=ft.Text(
                                value="Quantity",
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_600
                            )
                        ),

                        # Price column
                        ft.Container(
                            padding=ft.padding.only(top=15, bottom=15),
                            width=70,
                            alignment=ft.alignment.center,
                            bgcolor=ft.Colors.GREY_200,
                            content=ft.Text(
                                value="Price",
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_600
                            )
                        )
                    ]
                )
            ]
        )

        self.spacing = 0
        self.controls = [
            self.app_bar,

            self.quotation_table
        ]

        if self.page.client_storage.contains_key("quotation"):
            for i in range(len(self.page.client_storage.get("quotation"))):
                # append data row content
                self.quotation_table.controls.append(
                    ft.Row(
                        expand=True,
                        spacing=1,
                        controls=[
                            # Sr No column content
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=60,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_100,
                                content=ft.Text(
                                    value=self.page.client_storage.get("quotation")[i][0],
                                    color=ft.Colors.BLACK,
                                )
                            ),

                            # Goods Description column content
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10, left=10, right=10),
                                width=200,
                                alignment=ft.alignment.center_left,
                                bgcolor=ft.Colors.GREY_100,
                                content=ft.Text(
                                    value=self.page.client_storage.get("quotation")[i][1],
                                    color=ft.Colors.BLACK
                                )
                            ),

                            # Quantity column content
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=80,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_100,
                                content=ft.Text(
                                    value=self.page.client_storage.get("quotation")[i][2],
                                    color=ft.Colors.BLACK
                                )
                            ),

                            # Price column
                            ft.Container(
                                padding=ft.padding.only(top=10, bottom=10),
                                width=70,
                                alignment=ft.alignment.center,
                                bgcolor=ft.Colors.GREY_100,
                                content=ft.Text(
                                    value=self.page.client_storage.get("quotation")[i][3],
                                    color=ft.Colors.BLACK
                                )
                            )
                        ]
                    )
                )

                # append data row border
                self.quotation_table.controls.append(
                    ft.Row(
                        expand=True,
                        spacing=1,
                        controls=[
                            ft.Container(
                                expand=True,
                                height=1,
                                bgcolor=ft.Colors.GREY_200
                            )
                        ]
                    )
                )

    # clear quotation
    def clear_quotation(self, e):
        self.page.client_storage.remove("quotation")

        self.snack_bar.content.value = "Quotation cleared"
        self.page.open(self.snack_bar)

        time.sleep(1)

        self.page.go("/products")
