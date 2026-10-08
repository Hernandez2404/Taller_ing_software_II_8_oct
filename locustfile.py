import random
import uuid
from locust import HttpUser, task, between

class LaravelApiUser(HttpUser):
    wait_time = between(0.5, 2.0)

    @task(5)
    def test_list_users(self):
        """GET /api/users con paginación aleatoria"""
        page = random.randint(1, 100)
        per_page = 50
        with self.client.get(
            f"/api/users?page={page}&per_page={per_page}",
            name="/api/users",
            catch_response=True
        ) as response:
            if response.status_code == 200:
                try:
                    data = response.json()
                    if "data" in data and isinstance(data["data"], list):
                        response.success()
                    else:
                        response.failure("Estructura JSON inválida: falta el campo 'data'")
                except Exception as e:
                    response.failure(f"Error al procesar JSON: {e}")
            else:
                response.failure(f"Código HTTP inesperado: {response.status_code}")

    @task(3)
    def test_list_emails(self):
        """GET /api/users/emails con paginación aleatoria"""
        page = random.randint(1, 100)
        per_page = 50
        with self.client.get(
            f"/api/users/emails?page={page}&per_page={per_page}",
            name="/api/users/emails",
            catch_response=True
        ) as response:
            if response.status_code == 200:
                try:
                    data = response.json()
                    if "data" in data and isinstance(data["data"], list):
                        response.success()
                    else:
                        response.failure("Estructura JSON inválida: falta el campo 'data'")
                except Exception as e:
                    response.failure(f"Error al procesar JSON: {e}")
            else:
                response.failure(f"Código HTTP inesperado: {response.status_code}")

    @task(3)
    def test_list_over_twenty(self):
        """GET /api/users/over-twenty con paginación aleatoria"""
        page = random.randint(1, 100)
        per_page = 50
        with self.client.get(
            f"/api/users/over-twenty?page={page}&per_page={per_page}",
            name="/api/users/over-twenty",
            catch_response=True
        ) as response:
            if response.status_code == 200:
                try:
                    data = response.json()
                    if "data" in data and isinstance(data["data"], list):
                        response.success()
                    else:
                        response.failure("Estructura JSON inválida: falta el campo 'data'")
                except Exception as e:
                    response.failure(f"Error al procesar JSON: {e}")
            else:
                response.failure(f"Código HTTP inesperado: {response.status_code}")

    @task(1)
    def test_bulk_create_users(self):
        """POST /api/users/bulk creando 3 usuarios con emails únicos dinámicos"""
        uid = uuid.uuid4().hex[:8]
        payload = {
            "users": [
                {
                    "name": f"User A {uid}",
                    "email": f"usera_{uid}_{random.randint(1000, 9999)}@test.com",
                    "birth_date": "1995-03-15",
                    "password": "password123"
                },
                {
                    "name": f"User B {uid}",
                    "email": f"userb_{uid}_{random.randint(1000, 9999)}@test.com",
                    "birth_date": "2000-07-22",
                    "password": "password123"
                },
                {
                    "name": f"User C {uid}",
                    "email": f"userc_{uid}_{random.randint(1000, 9999)}@test.com",
                    "birth_date": "1998-11-05",
                    "password": "password123"
                }
            ]
        }
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        with self.client.post(
            "/api/users/bulk",
            json=payload,
            headers=headers,
            name="/api/users/bulk",
            catch_response=True
        ) as response:
            if response.status_code == 201:
                try:
                    data = response.json()
                    if "users" in data and len(data["users"]) == 3:
                        response.success()
                    else:
                        response.failure("Respuesta 201 incompleta: no retornó los 3 usuarios")
                except Exception as e:
                    response.failure(f"Error al procesar JSON: {e}")
            elif response.status_code == 422:
                response.failure(f"Error de validación 422 (email duplicado): {response.text}")
            else:
                response.failure(f"Código HTTP inesperado: {response.status_code} - {response.text}")
