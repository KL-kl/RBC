"""Relay_forecasting_Tool URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.conf.urls import url, include
from django.contrib import admin
from user.views import index,relay

urlpatterns = [
    url(r'^admin/', admin.site.urls),
    url(r'^$', index, name='index'),
    url(r'^relay/', relay, name='relay'),
    url(r'^CMNETwinUpload/', include(('CMNETwinUpload.urls', 'CMNETwinUpload'), namespace='CMNETwinUpload')),
    url(r'^IPwinUpload/', include(('IPwinUpload.urls', 'IPwinUpload'), namespace='IPwinUpload')),
    url(r'^downloadModel/', include(('downloadModel.urls', 'downloadModel'), namespace='downloadModel')),
    url(r'^user/', include(('user.urls', 'user'), namespace='user')),

]
