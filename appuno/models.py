from django.db import models
from .db_connection import db
# Create your models here.

datos_collection = db['datos']
camiones_collection = db['camion']