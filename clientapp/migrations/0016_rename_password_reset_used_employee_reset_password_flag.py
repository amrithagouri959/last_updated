# Generated by Django 5.0.7 on 2024-10-07 04:47

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('clientapp', '0015_rename_has_reset_password_employee_password_reset_used'),
    ]

    operations = [
        migrations.RenameField(
            model_name='employee',
            old_name='password_reset_used',
            new_name='reset_password_flag',
        ),
    ]