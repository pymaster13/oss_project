"""OSS URL Configuration"""

from django.urls import path

from .views import *

urlpatterns = [
    path('object_add/', object_add, name='object_add'),
    path('object_add_data_smeta/', object_add_data_smeta, name='object_add_data_smeta'),
    path('object_select_view/', object_select_view, name='object_select_view'),
    path('object_smeta/', object_smeta, name='object_smeta'),
    path('object_update/', object_update, name='object_update'),
    path('form_document/', form_document, name='form_document'),
    path('', index, name='index'),
]
