from django import forms

class PDFUploadForm(forms.Form):
    pdf_file = forms.FileField(widget=forms.FileInput(attrs={"id":"pdf_file","allow_multiple_selected": True}))
