from django.http import FileResponse, HttpResponse
from django.shortcuts import render

from converter.doc_converter import process_files

from .forms import ConversionData

def submit(request):
	print("t1")
	return index(request)

def index(request):
	print(f"got {request.method}")
	if request.method == "POST":
		res = FileResponse(process_files(request.FILES.getlist('powerpoints')))
		res['Content-Disposition'] = 'attachment; filename="converted.txt"'
		return res

	form = ConversionData()
	context = {"form": form}

	return render(request, "converter/index.html", context)
