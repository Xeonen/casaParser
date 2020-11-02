from django.shortcuts import render
from django.contrib.auth.models import User

# Create your views here.


def authPage(request):
    if request.method == "GET":
        return(render(request, "docParser/auth.html"))
    else:
        return (render(request, "docParser/auth.html"))

def cParser(request):
    return(render(request, "docParser/cParser.html"))