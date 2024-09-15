from django.urls import path
from .views import FileUploadView

urlpatterns = [
    path('uploadbordgais/', FileUploadView.as_view(), name='file-upload'),
]
