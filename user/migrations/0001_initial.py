# Generated by Django 4.1.1 on 2022-09-19 09:39

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='User',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nom', models.CharField(max_length=30, null=True)),
                ('prenom', models.CharField(max_length=30, null=True)),
                ('age', models.IntegerField(null=True)),
                ('adresse', models.CharField(max_length=80, null=True)),
            ],
        ),
    ]
