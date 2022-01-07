package ws.flota;

import java.io.Serializable;
import javax.jws.WebService;
import javax.jws.WebMethod;
import javax.jws.WebParam;

@WebService(serviceName = "ServiciosFlota")
public class ServiciosFlota implements Serializable {
    
    private static final long serialVersionUID = 1L;

    @WebMethod(operationName = "cargar_archivo")
    public String cargar_archivo(@WebParam(name = "path_archivo") String path_archivo) {
        String resultado;

        try {
            Driver driver = new Driver();
            driver.openConn();
            driver.SetAutoCommitConn(false);
            
            System.out.println("************ LIMPIA PAPELERA DE RECIBLAJE.");
            driver.Reciclebin();
            driver.commitConn();
            
            System.out.println("************ VALIDAR HOJAS DEL ARCHIVO EXCEL.");
            resultado = driver.validar_hojas_excel(path_archivo);
            
            if(resultado.equals("EXITO")) {
                System.out.println("************ INICIANDO CARGA VEHICULO.");
                resultado = driver.cargar_vehiculo(path_archivo);
            }
            if(resultado.equals("EXITO")) {
                System.out.println("************ INICIANDO CARGA VEHICULO TARJETA.");
                resultado = driver.cargar_vehiculo_tarjeta(path_archivo);
            }
            if(resultado.equals("EXITO")) {
                System.out.println("************ INICIANDO CARGA ESTACION.");
                resultado = driver.cargar_estacion(path_archivo);
            }
            if(resultado.equals("EXITO")) {
                System.out.println("************ INICIANDO CARGA PRODUCTO.");
                resultado = driver.cargar_producto(path_archivo);
            }
            if(resultado.equals("EXITO")) {
                System.out.println("************ INICIANDO CARGA HORARIO.");
                resultado = driver.cargar_horario(path_archivo);
            }
            if(resultado.equals("EXITO")) {
                System.out.println("************ INICIANDO CARGA CANTIDAD.");
                resultado = driver.actualiza_cantidad(path_archivo);
            }
            
            if(resultado.equals("EXITO")) {
                resultado = resultado + ":" + driver.getNumero_vehiculos();
                driver.commitConn();
            } else {
                driver.rollbackConn();
            }
            
            driver.SetAutoCommitConn(true);
            driver.closeConn();
            
        } catch (Exception ex) {
            resultado = "SERVICIO ERROR: " + ex.toString();
        }

        return resultado;
    }

}
