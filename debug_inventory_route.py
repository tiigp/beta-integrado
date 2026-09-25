import sys
sys.path.insert(0, 'C:/Users/PC/OneDrive/Desktop/BETA INTEGRADO/Control de combustible')
from app import app

with app.test_client() as c:
    r = c.get('/inventario')
    print('status=', r.status_code)
    print(r.data[:1000].decode('utf-8', 'ignore'))
