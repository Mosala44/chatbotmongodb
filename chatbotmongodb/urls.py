"""chatbotmongodb URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.1/topics/http/urls/
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
from django.contrib import admin
from django.urls import path
from appuno.views import lista_camiones, create_camion, Chatbot, index, selecciona_camion, generar_informe

urlpatterns = [
    path('admin/', admin.site.urls),
    path('listar_camiones/',lista_camiones, name='lista_camiones'),  # ✅ Nombre correcto
    path('crear_camion/', create_camion, name='create_camion'),
    path("chatbot/", Chatbot, name="Chatbot"),
    path('cht/', index, name= "cht"),
    path('', selecciona_camion, name='selecciona_camion'),
    path('descargar/', generar_informe, name= 'generar_informe')
     

]
