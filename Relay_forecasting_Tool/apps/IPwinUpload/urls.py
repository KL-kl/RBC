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
from IPwinUpload.views import IPwin, IPwin_2, upload_Ifile, up_maxfile, up_metric1, up_metric2, up_metric3, up_metric4, \
    down_metric1, ip_gz, te_upload, te_sortPath,search_routefile, iprelay_calculation,down_route,report_down_load,iprelay_compare

urlpatterns = [
    url(r'^IPwin/$', IPwin, name='IPwin'),
    url(r'^IPwin_2/$', IPwin_2, name='IPwin_2'),
    url(r'^upload_Ifile/$', upload_Ifile, name='upload_Ifile'),
    url(r'^report_down_load/$', report_down_load, name='report_down_load'),
    url(r'^up_maxfile/$', up_maxfile, name='up_maxfile'),
    url(r'^up_metric1/$', up_metric1, name='up_metric1'),
    url(r'^up_metric2/$', up_metric2, name='up_metric2'),
    url(r'^up_metric3/$', up_metric3, name='up_metric3'),
    url(r'^up_metric4/$', up_metric4, name='up_metric4'),
    url(r'^down_metric1/$', down_metric1, name='down_metric1'),
    url(r'^ip_gz/$', ip_gz, name='ip_gz'),
    url(r'^te_upload/$', te_upload, name='te_upload'),
    url(r'^te_sortPath/$', te_sortPath, name='te_sortPath'),
    url(r'^search_routefile/$', search_routefile, name='search_routefile'),
    url(r'^iprelay_calculation/$', iprelay_calculation, name='iprelay_calculation'),
url(r'^down_route/$', down_route, name='down_route'),
url(r'^iprelay_compare/$', iprelay_compare, name='iprelay_compare'),
]
