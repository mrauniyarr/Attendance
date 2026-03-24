import flet as ft
import datetime
import os
import shutil
import excel_handler as eh

async def main(page: ft.Page):
    page.title = "Attendance App"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 0
    page.window_width = 400
    page.window_height = 800
    
    # State
    selected_date = datetime.date.today().strftime("%d-%m-%Y")
    selected_class = None
    roll_number_input = ft.Text("", size=30, weight=ft.FontWeight.BOLD)
    
    # Share service
    share_service = ft.Share()
    
    # --- UI Helpers ---
    def show_snackbar(message, color=ft.Colors.GREEN):
        page.snack_bar = ft.SnackBar(ft.Text(message), bgcolor=color)
        page.snack_bar.open = True
        page.update()

    # --- Settings Tab Logic ---
    def handle_date_change(e):
        nonlocal selected_date
        selected_date = date_picker.value.strftime("%d-%m-%Y")
        date_btn.content = f"Date: {selected_date}"
        page.update()

    def add_new_class(e):
        if class_name_field.value:
            if eh.add_class(class_name_field.value):
                show_snackbar(f"Class '{class_name_field.value}' added!")
                refresh_dropdowns()
                class_name_field.value = ""
            else:
                show_snackbar("Class already exists!", ft.Colors.RED)
        page.update()

    def reset_all_data(e):
        def confirm_reset(e):
            eh.reset_data()
            refresh_dropdowns()
            show_snackbar("All data has been reset.")
            confirm_dialog.open = False
            page.update()

        confirm_dialog = ft.AlertDialog(
            title=ft.Text("Reset All Data?"),
            content=ft.Text("This will delete the entire Excel file. Are you sure?"),
            actions=[
                ft.TextButton("Cancel", on_click=lambda _: setattr(confirm_dialog, 'open', False) or page.update()),
                ft.Button(content="Reset", color=ft.Colors.WHITE, bgcolor=ft.Colors.RED, on_click=confirm_reset),
            ]
        )
        page.show_dialog(confirm_dialog)

    async def share_excel(e):
        if not os.path.exists(eh.EXCEL_FILE):
            show_snackbar("No data to export!", ft.Colors.RED)
            return

        try:
            if page.platform in (ft.PagePlatform.WINDOWS, ft.PagePlatform.LINUX):
                dest = os.path.join(
                    os.path.expanduser("~"),
                    f"Attendance_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                )
                shutil.copy2(eh.EXCEL_FILE, dest)
                show_snackbar("Saved to Downloads or user directory!", ft.Colors.BLUE)
            else:  # Android / iOS
                sp = ft.StoragePaths()
                temp_dir = await sp.get_temporary_directory()

                export_name = f"Attendance_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                temp_file = os.path.join(temp_dir, export_name)

                shutil.copy2(eh.EXCEL_FILE, temp_file)

                result = await share_service.share_files(
                    [ft.ShareFile.from_path(temp_file)], text="Master Attendance Report"
                )
                show_snackbar(f"Share sheet opened! ({result.status})", ft.Colors.GREEN)

        except Exception as ex:
            show_snackbar(f"Share error: {str(ex)[:80]}", ft.Colors.RED)
        page.update()

    def calc_percentage(e):
        if not selected_class:
            show_snackbar("Select a class first!", ft.Colors.RED)
            return
        msg = eh.calculate_percentage(selected_class)
        show_snackbar(msg)

    def del_percentage(e):
        if not selected_class:
            show_snackbar("Select a class first!", ft.Colors.RED)
            return
        msg = eh.delete_percentage_column(selected_class)
        show_snackbar(msg)

    # --- Attendance Tab Logic ---
    def keypad_click(e):
        val = e.control.content
        roll_number_input.value += val
        page.update()

    def clear_input(e):
        roll_number_input.value = ""
        page.update()

    def save_attendance(e):
        if not selected_class:
            show_snackbar("Select a class in Settings first!", ft.Colors.RED)
            return
        if not roll_number_input.value:
            show_snackbar("Enter a roll number!", ft.Colors.ORANGE)
            return
        
        eh.add_attendance(selected_class, selected_date, roll_number_input.value)
        show_snackbar(f"Roll {roll_number_input.value} marked Present!")
        roll_number_input.value = ""
        page.update()

    def delete_entry(e):
        if not selected_class:
            show_snackbar("Select a class in Settings first!", ft.Colors.RED)
            return
        if not roll_number_input.value:
            show_snackbar("Enter a roll number to delete!", ft.Colors.ORANGE)
            return
        
        deleted = eh.delete_attendance(selected_class, selected_date, roll_number_input.value)
        if deleted:
            show_snackbar(f"Deleted entry for Roll {roll_number_input.value}")
        else:
            show_snackbar(f"Roll {roll_number_input.value} had no attendance on this date.", ft.Colors.ORANGE)
        roll_number_input.value = ""
        page.update()

    # --- Statistics Tab Logic ---
    def view_stats(e):
        if not stats_class_dropdown.value:
            show_snackbar("Select a class!", ft.Colors.RED)
            return
        if not stats_roll_field.value:
            show_snackbar("Enter roll number!", ft.Colors.RED)
            return
        
        res = eh.get_student_stats(stats_class_dropdown.value, stats_roll_field.value)
        if res:
            color = ft.Colors.GREEN if float(res['percentage'][:-1]) >= 75 else ft.Colors.RED
            stats_result_card.content = ft.Container(
                padding=20,
                content=ft.Column([
                    ft.Text(f"Roll number: {stats_roll_field.value}", size=20, weight="bold"),
                    ft.Divider(),
                    ft.Row([ft.Text("Attendance:"), ft.Text(f"{res['days']} / {res['total']} days", weight="bold")]),
                    ft.Row([
                        ft.Text("Percentage:"),
                        ft.Text(res['percentage'], size=22, color=color, weight="bold")
                    ])
                ])
            )
            stats_result_card.visible = True
        else:
            show_snackbar("Student not found in this class!", ft.Colors.RED)
            stats_result_card.visible = False
        page.update()

    def remove_class_action(e):
        nonlocal selected_class
        if class_dropdown.value:
            if eh.remove_class(class_dropdown.value):
                show_snackbar(f"Class '{class_dropdown.value}' removed!")
                refresh_dropdowns()
                selected_class = None
                class_dropdown.value = None
            else:
                show_snackbar("Failed to remove class.", ft.Colors.RED)
        else:
            show_snackbar("Select a class to remove first!", ft.Colors.ORANGE)
        page.update()

    # --- UI Elements ---
    date_picker = ft.DatePicker(on_change=handle_date_change)
    page.overlay.append(date_picker)

    def open_date_picker(_):
        date_picker.open = True
        page.update()
    
    date_btn = ft.Button(
        content=f"Date: {selected_date}",
        icon=ft.Icons.CALENDAR_MONTH,
        on_click=open_date_picker,
        expand=True
    )
    
    class_dropdown = ft.Dropdown(label="Select Class", expand=True)
    def on_class_select(e):
        nonlocal selected_class
        selected_class = e.control.value
        page.update()
        
    class_dropdown.on_select = on_class_select
    
    class_name_field = ft.TextField(label="New Class Name", expand=True)
    
    # Settings Layout
    settings_view = ft.Container(
        padding=20,
        content=ft.Column([
            ft.Text("Application Settings", size=24, weight="bold"),
            # Part 1: Date and Class Management
            ft.Card(
                content=ft.Container(
                    padding=15,
                    content=ft.Column([
                        ft.Text("Date & Class Management", weight="bold", size=18),
                        ft.Row([date_btn]),
                        ft.Row([class_dropdown, ft.IconButton(ft.Icons.DELETE_OUTLINE, tooltip="Remove Selected Class", icon_color=ft.Colors.RED, on_click=remove_class_action)]),
                        ft.Row([class_name_field, ft.IconButton(ft.Icons.ADD_CIRCLE, tooltip="Add New Class", on_click=add_new_class, icon_color=ft.Colors.BLUE_ACCENT)]),
                    ], spacing=10)
                )
            ),
            # Part 2: Data Management
            ft.Card(
                content=ft.Container(
                    padding=15,
                    content=ft.Column([
                        ft.Text("Data Management", weight="bold", size=18),
                        ft.Row([
                            ft.Button(content="Add %", icon=ft.Icons.PERCENT, on_click=calc_percentage, expand=True,bgcolor=ft.Colors.GREEN,color=ft.Colors.WHITE),
                            ft.Button(content="Delete %", icon=ft.Icons.DELETE_SWEEP, on_click=del_percentage, expand=True,bgcolor=ft.Colors.BLUE,color=ft.Colors.WHITE),
                        ]),
                        ft.Row([
                            ft.Button(content="Share Excel", icon=ft.Icons.SHARE, on_click=share_excel, expand=True, bgcolor=ft.Colors.BLUE, color=ft.Colors.WHITE),
                            ft.Button(content="Reset All", icon=ft.Icons.RESTART_ALT, on_click=reset_all_data, expand=True, bgcolor=ft.Colors.RED, color=ft.Colors.WHITE),
                        ]),
                    ], spacing=10)
                )
            )
        ], spacing=15, scroll=ft.ScrollMode.AUTO)
    )

    # Attendance Layout
    keypad_buttons = []
    for i in range(1, 10):
        keypad_buttons.append(ft.Button(content=str(i), on_click=keypad_click, width=80, height=60, style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10))))
    keypad_buttons.append(ft.Button(content="C", on_click=clear_input, width=80, height=60, bgcolor=ft.Colors.GREY_300))
    keypad_buttons.append(ft.Button(content="0", on_click=keypad_click, width=80, height=60))
    keypad_buttons.append(ft.Container(width=80)) # Placeholder

    attendance_view = ft.Container(
        padding=20,
        content=ft.Column([
            ft.Text("Mark Attendance", size=24, weight="bold"),
            ft.Container(
                content=ft.Column([
                    ft.Text("Enter Roll Number", size=14, color=ft.Colors.GREY_700),
                    ft.Container(
                        content=roll_number_input,
                        bgcolor=ft.Colors.BLUE_50,
                        padding=15,
                        border_radius=10,
                        alignment=ft.Alignment(0, 0)
                    ),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                margin=ft.Margin.only(bottom=20)
            ),
            # Keypad
            ft.Column([
                ft.Row(keypad_buttons[0:3], alignment=ft.MainAxisAlignment.CENTER),
                ft.Row(keypad_buttons[3:6], alignment=ft.MainAxisAlignment.CENTER),
                ft.Row(keypad_buttons[6:9], alignment=ft.MainAxisAlignment.CENTER),
                ft.Row(keypad_buttons[9:12], alignment=ft.MainAxisAlignment.CENTER),
            ], spacing=10),
            ft.Divider(),
            ft.Row([
                ft.Button(content="SAVE", icon=ft.Icons.SAVE, on_click=save_attendance, expand=1, height=50, bgcolor=ft.Colors.GREEN, color=ft.Colors.WHITE),
                ft.Button(content="DELETE", icon=ft.Icons.DELETE, on_click=delete_entry, expand=1, height=50, bgcolor=ft.Colors.RED_ACCENT, color=ft.Colors.WHITE),
            ], spacing=10)
        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER)
    )

    # Statistics Layout
    stats_class_dropdown = ft.Dropdown(label="Select Class", expand=True)
    stats_roll_field = ft.TextField(label="Roll Number", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
    stats_result_card = ft.Card(
        elevation=5,
        content=ft.Container(padding=20),
        visible=False
    )

    stats_view = ft.Container(
        padding=20,
        content=ft.Column([
            ft.Text("Statistics", size=24, weight="bold"),
            ft.Divider(),
            ft.Card(
                content=ft.Container(
                    padding=15,
                    content=ft.Column([
                        stats_class_dropdown,
                        stats_roll_field,
                        ft.Button(content="Check Stats", icon=ft.Icons.ANALYTICS, on_click=view_stats, bgcolor=ft.Colors.BLUE_ACCENT, color=ft.Colors.WHITE),
                    ], spacing=15, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                )
            ),
            stats_result_card
        ], spacing=15, scroll=ft.ScrollMode.AUTO)
    )

    def refresh_dropdowns():
        classes = eh.get_all_classes()
        options = [ft.DropdownOption(c) for c in classes]
        class_dropdown.options = options
        stats_class_dropdown.options = options
        page.update()

    tabs = ft.Tabs(
        length=3,
        selected_index=1,
        animation_duration=300,
        content=ft.Column([
            ft.TabBar(
                tabs=[
                    ft.Tab(label="Settings", icon=ft.Icons.SETTINGS),
                    ft.Tab(label="Mark Attendance", icon=ft.Icons.CHECK_CIRCLE),
                    ft.Tab(label="Statistics", icon=ft.Icons.QUERY_STATS),
                ]
            ),
            ft.TabBarView(
                controls=[
                    settings_view,
                    attendance_view,
                    stats_view
                ],
                expand=1
            )
        ], expand=1),
        expand=1
    )

    page.add(tabs)
    refresh_dropdowns()

if __name__ == "__main__":
    ft.run(main)
