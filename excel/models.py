from django.db import models

class ExcelData(models.Model):
    sr_no = models.IntegerField()
    name = models.CharField(max_length=100)
    amount = models.DecimalField(max_digits=1000000000000, decimal_places=2)
    pending = models.BooleanField()

    def __str__(self):
        return f"Entry {self.sr_no}: {self.name}"
