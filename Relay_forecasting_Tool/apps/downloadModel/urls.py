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

from django.conf.urls import url
from downloadModel.views import downloadTest1,downloadTest2,downloadTest3,downloadTest4

urlpatterns = [
    url(r'^downloadTest1/$', downloadTest1, name='downloadTest1'),
    url(r'^downloadTest2/$', downloadTest2, name='downloadTest2'),
    url(r'^downloadTest3/$', downloadTest3, name='downloadTest3'),
    url(r'^downloadTest4/$', downloadTest4, name='downloadTest4'),

]
