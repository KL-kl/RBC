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
from CMNETwinUpload.views import CMNETwin, CMNETwin_2, up_metric1, up_metric2, up_metric3, up_metric4, \
    upload_Cfile, upload_newCfile, show_node, show_metric, show_trunkcircuit, \
    report_down_load, down_metric1, show_progress1, show_progress2, \
    up_flow, morenode, ntonflow, import_excel, gz, link_fail, search_routefile, down_route, relay_calculation, \
    relay_compare

urlpatterns = [
    url(r'^CMNETwin/$', CMNETwin, name='CMNETwin'),
    url(r'^CMNETwin_2/$', CMNETwin_2, name='CMNETwin_2'),
    url(r'^show_progress1/$', show_progress1, name='show_progress1'),
    url(r'^show_progress2/$', show_progress2, name='show_progress2'),
    url(r'^upload_Cfile/$', upload_Cfile, name='upload_Cfile'),
    url(r'^upload_newCfile/$', upload_newCfile, name='upload_newCfile'),
    url(r'^show_node/$', show_node, name='show_node'),
    url(r'^show_metric/$', show_metric, name='show_metric'),
    url(r'^show_trunkcircuit/$', show_trunkcircuit, name='show_trunkcircuit'),
    url(r'^down_metric1/$', down_metric1, name='down_metric1'),
    url(r'^report_down_load/$', report_down_load, name='report_down_load'),
    url(r'^up_flow/$', up_flow, name='up_flow'),
    url(r'^morenode/$', morenode, name='morenode'),
    url(r'^ntonflow/$', ntonflow, name='ntonflow'),
    url(r'^up_metric1/$', up_metric1, name='up_metric1'),
    url(r'^up_metric2/$', up_metric2, name='up_metric2'),
    url(r'^up_metric3/$', up_metric3, name='up_metric3'),
    url(r'^up_metric4/$', up_metric4, name='up_metric4'),
    url(r'^import_excel/$', import_excel, name='import_excel'),
    url(r'^gz/$', gz, name='gz'),
    url(r'^link_fail/$', link_fail, name='link_fail'),
    url(r'^search_routefile/$', search_routefile, name='search_routefile'),
    url(r'^down_route/$', down_route, name='down_route'),
    url(r'^relay_calculation/$', relay_calculation, name='relay_calculation'),
    url(r'^relay_compare/$', relay_compare, name='relay_compare'),

]
