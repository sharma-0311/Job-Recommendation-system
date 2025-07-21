import os
import pickle
import pandas as pd
from django.conf import settings
from django.shortcuts import render, get_object_or_404, redirect
import difflib
from .models import Job
from django.db.models import Count
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import UserCreationForm
from django.contrib.auth.forms import UserChangeForm
from .models import Profile
from django import forms
from django.contrib.auth.models import User

# Path to your models directory
MODELS_DIR = os.path.join(settings.BASE_DIR, 'static', 'models')

# Load the preprocessed jobs DataFrame
with open(os.path.join(MODELS_DIR, 'jobs.pkl'), 'rb') as f:
    jobs = pickle.load(f)

print("JOBS COLUMNS:", jobs.columns)

# Load the similarity matrix
with open(os.path.join(MODELS_DIR, 'similarity.pkl'), 'rb') as f:
    similarity = pickle.load(f)

def recommend(job_index, top_n=5):
    sim_scores = list(enumerate(similarity[job_index]))
    sim_scores = sorted(sim_scores, key=lambda x: x[1], reverse=True)
    sim_scores = sim_scores[1:top_n+1]
    recommended_indices = [i[0] for i in sim_scores]
    return jobs.iloc[recommended_indices]

def job_recommendation_view(request):
    recommendations = None
    error = None
    matched_title = None
    filters = {
        'Location': request.GET.get('location', ''),
        'Year': request.GET.get('year', ''),
        'Month': request.GET.get('month', ''),
        'IT': request.GET.get('it', ''),
    }
    locations = sorted(Job.objects.exclude(Location=None).values_list('Location', flat=True).distinct())
    years = sorted(Job.objects.exclude(Year=None).values_list('Year', flat=True).distinct())
    months = sorted(Job.objects.exclude(Month=None).values_list('Month', flat=True).distinct())
    its = ['True', 'False']

    # Load jobs DataFrame for index lookup only (not for display)
    import pickle
    import pandas as pd
    jobs_df = None
    with open(os.path.join(MODELS_DIR, 'jobs.pkl'), 'rb') as f:
        jobs_df = pickle.load(f)

    if request.method == 'POST':
        job_title = request.POST.get('job_title')
        all_titles = list(jobs_df['Title'])
        matches = difflib.get_close_matches(job_title, all_titles, n=1, cutoff=0.5)
        if matches:
            matched_title = matches[0]
            job_index = jobs_df[jobs_df['Title'] == matched_title].index[0]
            sim_scores = list(enumerate(similarity[job_index]))
            sim_scores = sorted(sim_scores, key=lambda x: x[1], reverse=True)
            sim_scores = sim_scores[1:6]  # Top 5 recommendations
            recommended_indices = [i[0] for i in sim_scores]
            jobs_qs = Job.objects.filter(id__in=[i+1 for i in recommended_indices])
            # Apply filters if needed (optional)
            recommendations = list(jobs_qs.values())
        else:
            error = 'Job not found. Please enter a valid job title.'
            recommendations = None
    return render(request, 'index.html', {
        'job_titles': list(Job.objects.values_list('Title', flat=True).distinct()),
        'recommendations': recommendations,
        'error': error,
        'matched_title': matched_title,
        'filters': filters,
        'locations': locations,
        'years': years,
        'months': months,
        'its': its,
        'show_all': False,
        'page': 1,
        'total_pages': 1,
    })

@login_required
def all_jobs_view(request):
    page = int(request.GET.get('page', 1))
    page_size = 30
    filters = {
        'Location': request.GET.get('location', ''),
        'Year': request.GET.get('year', ''),
        'Month': request.GET.get('month', ''),
        'IT': request.GET.get('it', ''),
    }
    # Top 100 most common locations
    locations = list(
        Job.objects.values_list('Location', flat=True)
        .exclude(Location=None)
        .annotate(num=Count('Location'))
        .order_by('-num')[:100]
    )
    years = sorted(Job.objects.exclude(Year=None).values_list('Year', flat=True).distinct())
    months = sorted(Job.objects.exclude(Month=None).values_list('Month', flat=True).distinct())
    its = ['True', 'False']

    all_jobs = Job.objects.all()
    if filters['Location']:
        all_jobs = all_jobs.filter(Location=filters['Location'])
    if filters['Year']:
        all_jobs = all_jobs.filter(Year=int(filters['Year']))
    if filters['Month']:
        all_jobs = all_jobs.filter(Month=int(filters['Month']))
    if filters['IT']:
        all_jobs = all_jobs.filter(IT=(filters['IT'] == 'True'))

    total_jobs = all_jobs.count()
    start = (page - 1) * page_size
    end = start + page_size
    jobs_page = all_jobs[start:end]
    total_pages = (total_jobs + page_size - 1) // page_size
    return render(request, 'index.html', {
        'job_titles': list(Job.objects.values_list('Title', flat=True).distinct()),
        'all_jobs': jobs_page,
        'recommendations': None,
        'error': None,
        'matched_title': None,
        'filters': filters,
        'locations': locations,
        'years': years,
        'months': months,
        'its': its,
        'show_all': True,
        'page': page,
        'total_pages': total_pages,
    })

@login_required
def job_detail_view(request, job_id):
    job = get_object_or_404(Job, id=job_id)
    # Find the index in the DataFrame (assuming IDs are 1-based and match DataFrame order)
    import pickle
    import pandas as pd
    with open(os.path.join(MODELS_DIR, 'jobs.pkl'), 'rb') as f:
        jobs_df = pickle.load(f)
    # job_id - 1 is the DataFrame index if imported in order
    job_index = job_id - 1
    sim_scores = list(enumerate(similarity[job_index]))
    sim_scores = sorted(sim_scores, key=lambda x: x[1], reverse=True)
    sim_scores = sim_scores[1:6]  # Top 5 recommendations
    recommended_indices = [i[0] for i in sim_scores]
    recommended_jobs = Job.objects.filter(id__in=[i+1 for i in recommended_indices])
    return render(request, 'job_detail.html', {'job': job, 'recommended_jobs': recommended_jobs})

# For recommendations, we still need the DataFrame for similarity, but job details should come from the DB 

@login_required
def profile_view(request):
    profile, created = Profile.objects.get_or_create(user=request.user)
    return render(request, 'profile.html', {'user': request.user, 'profile': profile}) 

def signup_view(request):
    if request.method == 'POST':
        form = UserCreationForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('login')
    else:
        form = UserCreationForm()
    return render(request, 'signup.html', {'form': form}) 

class ProfileForm(forms.ModelForm):
    class Meta:
        model = Profile
        fields = ['photo', 'bio', 'phone']

class UserUpdateForm(forms.ModelForm):
    class Meta:
        model = User
        fields = ['first_name', 'last_name', 'email']

@login_required
def profile_update_view(request):
    profile, created = Profile.objects.get_or_create(user=request.user)
    if request.method == 'POST':
        user_form = UserUpdateForm(request.POST, instance=request.user)
        profile_form = ProfileForm(request.POST, request.FILES, instance=profile)
        if user_form.is_valid() and profile_form.is_valid():
            user_form.save()
            profile_form.save()
            return redirect('profile')
    else:
        user_form = UserUpdateForm(instance=request.user)
        profile_form = ProfileForm(instance=profile)
    return render(request, 'profile_update.html', {
        'user_form': user_form,
        'profile_form': profile_form,
        'profile': profile,
    }) 