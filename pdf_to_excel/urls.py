from django.urls import path
from . import views
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('', views.upload_pdf, name='upload_pdf'),
    path('download_file/<path:filename>/', views.download_file, name='download_file'),]
