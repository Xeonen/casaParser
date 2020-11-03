from django.shortcuts import render
from django.contrib.auth.models import User
from django.db import IntegrityError
from django.contrib.auth import login, logout, authenticate
from django.shortcuts import redirect
from django.contrib.auth.forms import AuthenticationForm
from .forms import UploadFileForm
from django.core.files.storage import FileSystemStorage
from zipfile import ZipFile
from django.http import HttpResponse
from django.views.static import serve
from shutil import rmtree


from .excelProcedure import excelProcedure
# Create your views here.


def authPage(request):
    if request.method == "GET":
        if request.user.is_authenticated:
            print(f"User is {request.user.is_authenticated}")
            return (redirect("cparser"))
        else:
            print(f"User is {request.user.is_authenticated}")
            return (render(request, "docParser/auth.html", {"form": AuthenticationForm}))
    else:
        user = authenticate(
            request,
            username=request.POST["username"],
            password=request.POST["password"]
        )
        if user is None:
            return (
                render(
                    request, "docParser/auth.html",
                    {"form": AuthenticationForm(), "error": "Kullanıcı adı veya şifresi hatalıdır."}
                    ))
        else:
            login(request, user)
            return(redirect("cparser"))

def logoutuser(request):
    if request.method == "POST":
        print("Logout Called")
        logout(request)
    return (redirect("auth"))

def cParser(request):
    try:
        rmtree("media")
    except:
        pass
    fileDict = {"casaFiles": "data.zip", "reportData": "dataset.xlsx", "sourceFile": "source.xlsx"}
    if request.method == "POST":
        try:
            formType = int(request.POST["formType"])
            payetSize = float(request.POST["payetSize"])
        except Exception as e:
            print(e)


        for file in request.FILES:
            uploadedFile = request.FILES[file]
            FileSystemStorage().save(fileDict[file], uploadedFile)


        with ZipFile("media/data.zip", "r") as zip:
            zip.extractall("media/data/")

        ep = excelProcedure("media/source.xlsx", "media/dataset.xlsx", formType, payetSize)
        parsedData = ep.fillForm()

        return(serve(request, "media/casaRapor.xlsx", ""))
    else:
        return (render(request, "docParser/cParser.html"))

# ep = excelProcedure("source.xlsx", "dataset.xlsx", 1, 0.25)
# ep.fillForm()
