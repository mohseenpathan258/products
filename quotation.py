import flet as ft
from home import HomePage


class QuotationPage(ft.Column):
    quotation_content = ft.Column(
        spacing=0,
    )

    def __init__(self, page, visibility_purchase_btn):
        super(QuotationPage, self).__init__()

        self.page = page

        self.visibility_purchase_btn = visibility_purchase_btn

        self.spacing = 0
        self.controls = [
            ft.Container(
                expand=True,
                bgcolor="#008FD5",
                margin=ft.margin.only(top=10),
                padding=10,
                alignment=ft.alignment.center,
                content=ft.Text(
                    value="Quotation",
                    color=ft.Colors.WHITE,
                    weight=ft.FontWeight.BOLD,
                    size=25
                )
            ),

            ft.Row(
                spacing=1,
                controls=[
                    ft.Container(
                        bgcolor=ft.Colors.GREY_200,
                        width=252,
                        height=40,
                        alignment=ft.alignment.center_left,
                        padding=ft.padding.only(left=10),
                        content=ft.Text(
                            value="Goods Description",
                            color=ft.Colors.BLACK,
                            weight=ft.FontWeight.W_500
                        )
                    ),

                    ft.Container(
                        bgcolor=ft.Colors.GREY_200,
                        width=60,
                        height=40,
                        alignment=ft.alignment.center,
                        content=ft.Text(
                            value="Pcs",
                            color=ft.Colors.BLACK,
                            weight=ft.FontWeight.W_500
                        )
                    ),

                    ft.Container(
                        bgcolor=ft.Colors.GREY_200,
                        width=100,
                        height=40,
                        alignment=ft.alignment.center,
                        content=ft.Text(
                            value="Price",
                            color=ft.Colors.BLACK,
                            weight=ft.FontWeight.W_500
                        )
                    ),
                ]
            ),

            self.quotation_content
        ]

        self.add_quotation_content()

    def add_quotation_content(self):
        # HomePage class instance
        home = HomePage(self.page, self.visibility_purchase_btn)

        # clear quotation_content
        self.quotation_content.controls.clear()

        # append quotation_content
        for i in home.quotation_list:
            self.quotation_content.controls.append(
                ft.Row(
                    spacing=1,
                    controls=[
                        # quotation Goods Description column
                        ft.Container(
                            bgcolor=ft.Colors.WHITE,
                            width=252,
                            padding=10,
                            alignment=ft.alignment.center_left,
                            content=ft.Text(
                                value=i[0],
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_500
                            )
                        ),

                        # quotation Pcs column
                        ft.Container(
                            bgcolor=ft.Colors.WHITE,
                            width=60,
                            alignment=ft.alignment.center,
                            content=ft.Text(
                                value=i[1],
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_500
                            )
                        ),

                        # quotation Price column
                        ft.Container(
                            bgcolor=ft.Colors.WHITE,
                            width=100,
                            alignment=ft.alignment.center,
                            content=ft.Text(
                                value=i[2],
                                color=ft.Colors.BLACK,
                                weight=ft.FontWeight.W_500
                            )
                        ),
                    ]
                )
            ),

            self.quotation_content.controls.append(
                ft.Row(
                    expand=True,
                    controls=[
                        ft.Container(
                            expand=True,
                            height=1,
                            bgcolor=ft.Colors.GREY_200
                        )
                    ]
                )
            )

        self.page.update()
