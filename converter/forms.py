from django import forms

class ConversionData(forms.Form):
	student_name = forms.CharField(label="student name", max_length=100)
	powerpoints = forms.FileField(widget = forms.TextInput(attrs={
            "name": "images",
            "type": "File",
            "class": "form-control",
            "multiple": "True",
        }), label = "")