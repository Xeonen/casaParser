from django.shortcuts import render
from django.contrib.auth.models import User
from django.db import IntegrityError
from django.contrib.auth import login, logout, authenticate
from django.shortcuts import redirect
from django.contrib.auth.forms import AuthenticationForm
from .forms import UploadFileForm

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
    return (render(request, "docParser/cParser.html"))
