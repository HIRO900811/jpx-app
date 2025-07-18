from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),  # トップページにアクセスしたら views.index 関数を実行
]