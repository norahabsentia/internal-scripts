from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField()

class UploadFileWithInputForm(forms.Form):
    link1 = forms.CharField()
    video = forms.CharField()
    videoLink = forms.CharField()
    endposter = forms.CharField()
    file = forms.FileField()
