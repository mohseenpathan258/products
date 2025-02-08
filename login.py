import flet as ft


class Login(ft.Column):
    def __init__(self, page):
        super(Login, self).__init__()

        self.page = page

        self.login_text_box = ft.CupertinoTextField(
            placeholder_text="Password",
        )

        self.controls = [
            ft.Container(
                height=100
            ),

            self.login_text_box,

            ft.CupertinoButton(
                text="Login",
                on_click=self.login
            )
        ]

    def login(self, e):
        if self.login_text_box.value == "mkp0258":
            self.page.go('/products')
