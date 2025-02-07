import flet as ft
from login import Login
from products import Products


def views_handler(page):
    return {
        '/login': ft.View(
            route='/login',
            padding=0,
            spacing=0,
            bgcolor=ft.Colors.WHITE,
            controls=[
                # login page
                Login(page)
            ]
        ),

        '/products': ft.View(
            route='/products',
            padding=0,
            spacing=0,
            bgcolor=ft.Colors.GREY_100,
            controls=[
                # products page
                Products(page)
            ]
        )
    }
