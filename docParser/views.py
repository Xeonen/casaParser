from django.shortcuts import render
from django.http import HttpResponse

# Create your views here.


def authPage(request):
    return(HttpResponse("Works"))