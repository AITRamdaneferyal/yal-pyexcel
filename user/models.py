from django.db import models
# Create your models here.
class User(models.Model):
    nom = models.CharField(max_length=30,null=True)
    prenom = models.CharField(max_length=30,null=True)
    age = models.IntegerField(null=True)
    adresse = models.CharField(max_length=80,null=True)


