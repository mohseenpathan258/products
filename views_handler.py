import flet as ft
from login import Login
from products import Products
from quotation import Quotation


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
        ),

        '/quotation': ft.View(
            route='/quotation',
            padding=0,
            spacing=0,
            scroll=ft.ScrollMode.HIDDEN,
            bgcolor=ft.Colors.GREY_100,
            controls=[
                # products page
                Quotation(page)
            ]
        )
    }
