from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload, name='upload'),
    path('process-upload/', views.process_upload, name='process_upload'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('generate-batch/', views.generate_batch, name='generate_batch'),
    path('display-data/', views.display_data, name='display_data'),
    path('display-data-table/', views.display_data_table, name='display_data_table'),
    path('download-excel/', views.download_excel, name='download_excel'),
    path('generate-final-batch/', views.generate_final_batch, name='generate_final_batch'),
    path('generate-commission-batch/', views.generate_commission_batch, name='generate_commission_batch'),
]