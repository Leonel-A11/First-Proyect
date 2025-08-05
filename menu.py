import funcionesexcel
import openpyxl

while True:
    opcion = int(input("Seleccione una opcion : \n 1-Registrar cliente \n 2-Registrar producto \n 3-Registrar Venta \n 4-Ver inventario \n 5-Mostrar registro de ventas \n 6 Actualizar producto \n 7 Verificar Productos \n 8-Salir \n : "))

    if opcion == 1:
        funcionesexcel.registrarcliente()
    elif opcion == 2:
        funcionesexcel.registrarproductos()
    elif opcion == 3:
        funcionesexcel.venta()
    elif opcion == 4:
        funcionesexcel.mostrarinventario()
    elif opcion == 5:
        funcionesexcel.mostrarregistroventas()
    elif opcion == 6:
        funcionesexcel.actualizarproducto()
    elif opcion == 7:
        funcionesexcel.ver_productos5()
    elif opcion == 8:
        break
    else:
        print ("Opcion no existe")
