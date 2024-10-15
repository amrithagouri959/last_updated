from django.contrib import admin
from .models import Employee, WorkSession
from django.http import HttpResponse
from django.urls import path
from openpyxl import Workbook
from io import BytesIO
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph

class WorkSessionAdmin(admin.ModelAdmin):
    list_display = ('employee', 'start_time', 'end_time', 'duration', 'machine', 'complaint', 'issue')
    list_filter = ('employee', 'start_time', 'end_time')
    search_fields = ('employee__name', 'start_time', 'end_time', 'complaint', 'issue')

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('export/excel/', self.admin_site.admin_view(self.export_as_excel), name='worksession-export-excel'),
            path('export/pdf/', self.admin_site.admin_view(self.export_as_pdf), name='worksession-export-pdf'),
        ]
        return custom_urls + urls

    # Export to Excel
    def export_as_excel(self, request):
        """Export work sessions to an Excel file."""
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=work_sessions.xlsx'

        wb = Workbook()
        ws = wb.active
        ws.title = "Work Sessions"

        # Define headers
        ws.append(['Employee', 'Start Time', 'End Time', 'Duration', 'Machine', 'Complaint', 'Issue'])

        # Add data
        for session in WorkSession.objects.all():
            # Remove timezone information from datetime fields
            row = [
                session.employee.name if session.employee else 'N/A',
                session.start_time.replace(tzinfo=None).isoformat() if session.start_time else 'N/A',
                session.end_time.replace(tzinfo=None).isoformat() if session.end_time else 'N/A',
                session.duration() if session.duration() else 'N/A',  # Ensure duration is not None
                session.machine if session.machine else 'N/A',
                session.complaint if session.complaint else 'N/A',
                session.issue if session.issue else 'N/A',
            ]
            ws.append(row)

        # Set column widths for better visibility
        column_widths = [25, 25, 25, 15, 20, 20, 30]  # Adjust as needed
        for i, width in enumerate(column_widths, 1):  # Excel columns are 1-indexed
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width

        wb.save(response)
        return response

    # Export to PDF
    def export_as_pdf(self, request):
        """Export work sessions to a PDF file."""
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="work_sessions.pdf"'

        buffer = BytesIO()
        # Use A3 size in landscape orientation for more width
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []

        # Title
        styles = getSampleStyleSheet()
        title = Paragraph("Work Session Report", styles['Title'])
        elements.append(title)

        # Table data
        data = [['Employee', 'Start Time', 'End Time', 'Duration', 'Machine', 'Complaint', 'Issue']]
        
        # Add data
        for session in WorkSession.objects.all():
            data.append([
                Paragraph(session.employee.name if session.employee else 'N/A', styles['Normal']),
                Paragraph(str(session.start_time) if session.start_time else 'N/A', styles['Normal']),
                Paragraph(str(session.end_time) if session.end_time else 'N/A', styles['Normal']),
                Paragraph(str(session.duration()) if session.duration() else 'N/A', styles['Normal']),
                Paragraph(session.machine if session.machine else 'N/A', styles['Normal']),
                Paragraph(session.complaint if session.complaint else 'N/A', styles['Normal']),
                Paragraph(session.issue if session.issue else 'N/A', styles['Normal']),
            ])

        # Create table and style with custom column widths
        table = Table(data, colWidths=[100, 120, 120, 80, 100, 100, 150])  # Adjust widths as necessary
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
        ]))

        elements.append(table)
        doc.build(elements)

        pdf = buffer.getvalue()
        buffer.close()
        response.write(pdf)
        return response

admin.site.register(Employee)
admin.site.register(WorkSession, WorkSessionAdmin)
