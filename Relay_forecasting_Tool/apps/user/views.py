from django.shortcuts import render, redirect,reverse
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from .forms import LoginForm

# Create your views here.
@login_required(login_url='/user/user_login/')
def index(request):
    return render(request, 'index.html')

@login_required(login_url='/user/user_login/')
def relay(request):
    return render(request, 'base.html')

# @login_required(login_url='/user/user_login/')
# def device(request):
#     return render(request, 'device.html')


def user_login(request):
    if request.method == 'POST':

        # 用户名密码验证
        user_login_form = LoginForm(request.POST)
        # username = request.POST.get('username')
        # password = request.POST.get('password')
        if user_login_form.is_valid():
            username = user_login_form.cleaned_data['username']
            password = user_login_form.cleaned_data['password']
            user = authenticate(username=username, password=password)

            if user:
                login(request, user)
                return redirect('/')
            else:
                return render(request,'login.html',{
                    'msg':'用户名或密码有误'
                })
        else:
            print(user_login_form)
            return render(request, 'login.html', {
                'user_login_form': user_login_form,

            })
    else:
        return render(request, 'login.html')


def user_logout(request):
    logout(request)
    return redirect('/user/user_login/')
