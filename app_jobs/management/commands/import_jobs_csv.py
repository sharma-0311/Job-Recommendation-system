from django.core.management.base import BaseCommand
import pandas as pd
from app_jobs.models import Job

class Command(BaseCommand):
    help = 'Import jobs from CSV into the Job model'

    def handle(self, *args, **kwargs):
        df = pd.read_csv('data job posts.csv')
        count = 0
        for _, row in df.iterrows():
            Job.objects.create(
                jobpost=row.get('jobpost', ''),
                date=row.get('date', ''),
                Title=row.get('Title', ''),
                Company=row.get('Company', ''),
                Location=row.get('Location', ''),
                JobDescription=row.get('JobDescription', ''),
                JobRequirment=row.get('JobRequirment', ''),
                RequiredQual=row.get('RequiredQual', ''),
                Year=row.get('Year', None),
                Month=row.get('Month', None),
                IT=bool(row.get('IT', False)),
                tags=row.get('tags', ''),
            )
            count += 1
        self.stdout.write(self.style.SUCCESS(f'Successfully imported {count} jobs from CSV.')) 