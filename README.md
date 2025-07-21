# AKS Job Listing - Django Job Recommendation System

## Overview

**AKS Job Listing** is a modern Django web application for job search and recommendations. It features:
- User authentication (signup, login, logout, profile with photo upload)
- Job listing with filters and pagination
- Job detail pages with sidebar recommendations (using ML similarity)
- Search bar for instant job recommendations
- Admin and user-friendly UI

---

## Features

- **User Authentication**: Signup, login, logout, and profile management (with photo upload and editable details)
- **Job Listing**: Paginated, filterable job list (Location, Year, Month, IT)
- **Job Detail**: Full job info with a sidebar of recommended jobs (ML-powered)
- **Search**: Search bar for job title, returns top 5 similar jobs
- **Admin**: Manage jobs and users via Django admin

---

## Setup Instructions

### 1. Clone the Repository
```bash
# Clone your repo (replace with your actual repo URL)
git clone <your-repo-url>
cd <your-repo-directory>
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
# Or manually:
pip install django pandas Pillow
```

### 3. Database & Migrations
```bash
python manage.py makemigrations
python manage.py migrate
```

### 4. Import Job Data
- Place your job data CSV (e.g., `data job posts.csv`) in the project root.
- Run:
```bash
python manage.py import_jobs_csv
```

### 5. Create a Superuser (for admin access)
```bash
python manage.py createsuperuser
```

### 6. Run the Development Server
```bash
python manage.py runserver
```

### 7. Media Files (Profile Photos)
- Make sure `MEDIA_URL` and `MEDIA_ROOT` are set in `settings.py`:
  ```python
  MEDIA_URL = '/media/'
  MEDIA_ROOT = BASE_DIR / 'media'
  ```
- In `urls.py`:
  ```python
  from django.conf import settings
  from django.conf.urls.static import static
  if settings.DEBUG:
      urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
  ```

---

## Usage

- **Home Page**: Shows all jobs (requires login)
- **Filters**: Use sidebar to filter jobs by location, year, month, IT
- **Pagination**: Browse jobs 30 per page
- **Search Bar**: Enter a job title to get top 5 recommended jobs (ML similarity)
- **Job Detail**: Click a job to see full details and sidebar recommendations
- **Profile**: Click the profile photo (top right) to view or edit your profile
- **Profile Edit**: Upload a photo, update bio, phone, name, email
- **Signup/Login/Logout**: Use the navbar links

---

## ML Recommendation System
- Uses a precomputed similarity matrix (`similarity.pkl`) and job data (`jobs.pkl` or CSV)
- When searching or viewing a job, the app finds the most similar jobs using cosine similarity
- Recommendations are fetched from the database using the indices from the similarity matrix

---

## File Structure
- `app_jobs/` - Main Django app (models, views, urls, management commands)
- `job_project/` - Django project settings and root URLs
- `templates/` - All HTML templates
- `static/` - Static files (CSS, models, etc.)
- `media/` - Uploaded profile photos
- `requirements.txt` - Python dependencies

---

## Customization
- Add more fields to the `Profile` model as needed
- Add more filters or search options in the job list
- Style the UI in `templates/` and static CSS
- Extend the admin for more control

---

## Troubleshooting
- **Image upload not working?** Make sure Pillow is installed and media settings are correct
- **Logout not working?** Ensure the logout URL is `/logout/` and uses Django's `LogoutView`
- **Job recommendations not showing?** Check that `similarity.pkl` and job data are in sync with the database

---

