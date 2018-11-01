"""
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
from django.contrib import admin
from django.urls import path, include
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('script1/', views.script1, name='script1'),
    path('script2/', views.script2, name='script2'),
    path('script3/', views.script3, name='script3'),
    path('script4/', views.script4, name='script4'),
    path('script5/', views.script5, name='script5'),
    path('script6/', views.nameSplit, name='script6'),
    path('ds6/', views.namesplit_download, name='ds6'),
]
