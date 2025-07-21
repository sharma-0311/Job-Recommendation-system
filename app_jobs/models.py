from django.db import models
from django.contrib.auth.models import User

# Create your models here.

class Job(models.Model):
    jobpost = models.TextField(blank=True, null=True)
    date = models.CharField(max_length=100, blank=True, null=True)
    Title = models.CharField(max_length=255)
    Company = models.CharField(max_length=255, blank=True, null=True)
    Location = models.CharField(max_length=255, blank=True, null=True)
    JobDescription = models.TextField(blank=True, null=True)
    JobRequirment = models.TextField(blank=True, null=True)
    RequiredQual = models.TextField(blank=True, null=True)
    Year = models.IntegerField(blank=True, null=True)
    Month = models.IntegerField(blank=True, null=True)
    IT = models.BooleanField(default=False)
    tags = models.TextField(blank=True, null=True)

    def __str__(self):
        return self.Title

class Profile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    photo = models.ImageField(upload_to='profile_photos/', blank=True, null=True)
    bio = models.TextField(blank=True, null=True)
    phone = models.CharField(max_length=20, blank=True, null=True)

    def __str__(self):
        return f"Profile of {self.user.username}"
