from django.urls import path
from django.contrib.auth import views as auth_views
from . import views
from .views import job_recommendation_view, all_jobs_view, job_detail_view, profile_view, signup_view, profile_update_view

urlpatterns = [
    path('', all_jobs_view, name="all_jobs"),
    path('jobs/', all_jobs_view, name="all_jobs"),
    path('recommend/', job_recommendation_view, name="job_recommendation"),
    path('job/<int:job_id>/', job_detail_view, name="job_detail"),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='/'), name='logout'),
    path('profile/', profile_view, name='profile'),
    path('profile/update/', profile_update_view, name='profile_update'),
    path('signup/', signup_view, name='signup'),
]