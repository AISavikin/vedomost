from peewee import *

db = SqliteDatabase('database.db')


class BaseModel(Model):
    class Meta:
        database = db


class Student(BaseModel):
    id = AutoField(primary_key=True)
    name = CharField(null=False)
    added = DateTimeField()
    group = IntegerField()
    active = BooleanField(default=True)

    class Meta:
        db_table = 'students'
        order_by = 'name'


class Attendance(BaseModel):
    id = ForeignKeyField(Student)
    day = IntegerField()
    month = CharField()
    year = IntegerField()
    absent = CharField()

    class Meta:
        db_table = 'attendances'


def create_database():
    db.create_tables([Student, Attendance])


if __name__ == '__main__':
    create_database()
