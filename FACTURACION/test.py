from security import security

AppKey = security()
print(AppKey["cadena"])
print(AppKey["clave_cifrada"])
print(AppKey["clave"])