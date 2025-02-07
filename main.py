import flet as ft
from views_handler import views_handler


def main(page: ft.Page):
    def page_route(route):
        page.views.clear()
        page.views.append(
            views_handler(page)[page.route]
        )
        page.update()

    page.theme_mode = ft.ThemeMode.LIGHT
    page.on_route_change = page_route

    page.go('/login')


ft.app(target=main, assets_dir='assets')
