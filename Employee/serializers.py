from rest_framework import serializers
from . import models as app_models
from django.db import models as django_models



class TeamSerializer(serializers.ModelSerializer):
    class Meta:
        model = app_models.Team
        fields = '__all__'
class DepartmentSerializer(serializers.ModelSerializer):
    class Meta:
        model = app_models.Department
        fields = '__all__'

class TeamCreate(serializers.ModelSerializer):
    class Meta:
        model = app_models.Team
        exclude = ['updated_by', 'created_by']
class TeamUpdate(serializers.ModelSerializer):
    class Meta:
        model = app_models.Team
        exclude = ['created_by', 'updated_by']
        
class DepartmentCreate(serializers.ModelSerializer):
    class Meta:
        model = app_models.Department
        exclude = ['updated_by', 'created_by']
class DepartmentUpdate(serializers.ModelSerializer):
    class Meta:
        model = app_models.Department
        exclude = ['created_by', 'updated_by']


class RegistrationSerializer(serializers.ModelSerializer):
    username = serializers.CharField()   
    def _user(self, obj):
        request = self.context.get('request', None)
        if request and request.user:
            return request.user.username
        return ''
    class Meta:
        model = app_models.Employee
        exclude = ['last_login', 'is_superuser', 'is_staff', 'is_active', 'is_admin', 
                   'date_joined', 'user_permissions', 'groups']
        extra_kwargs = {'password': {'write_only': True}}

class EmployeeSerializer(serializers.ModelSerializer):
    class Meta:
        model = app_models.Employee
        fields = ['username', 'department', 'team', 'id', 'position', 'is_manager', 'manages', 'full_name']

class EmployeeSerializerExtended(serializers.ModelSerializer):
    
    class Meta:
        model = app_models.Employee
        fields = '__all__'
class LoginSerializer(serializers.Serializer):
    username = serializers.EmailField(write_only=True)
    password = serializers.CharField(write_only=True)

class LogoutSerializer(serializers.Serializer):
    refresh_token = serializers.CharField(write_only=True)


class exportExcelSerializer(serializers.Serializer):
    deadline = serializers.CharField(write_only=True)
    department_id= serializers.IntegerField(write_only=True)