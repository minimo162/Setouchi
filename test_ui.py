import flet as ft
import sys

def main(page: ft.Page):
    page.title = 'UI Test'
    page.bgcolor = '#FFFFFF'

    # Simple test
    page.add(ft.Text('DEBUG: page.add works!', size=30, color='red'))

    # Test with spans
    try:
        span1 = ft.TextSpan('Hello ', style=ft.TextStyle(size=20, color='blue'))
        span2 = ft.TextSpan('World', style=ft.TextStyle(size=20, weight=ft.FontWeight.BOLD))
        text_with_spans = ft.Text(spans=[span1, span2])
        page.add(text_with_spans)
        print('Text with spans added successfully')
    except Exception as e:
        print(f'Error with spans: {e}')
        page.add(ft.Text(f'Spans error: {e}', color='red'))

    print('UI setup complete')

if __name__ == '__main__':
    print(f'Flet version: {ft.version.version}')
    ft.app(target=main, view=None, port=8551)
