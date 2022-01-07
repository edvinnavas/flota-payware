package com.flota;

import java.io.Serializable;

public class Servicio implements Serializable {
    
    private static final long serialVersionUID = 1L;
    
    public Servicio() {
        
    }

    public String cargarArchivo(java.lang.String pathArchivo) {
        ws.flota.ServiciosFlota_Service service = new ws.flota.ServiciosFlota_Service();
        ws.flota.ServiciosFlota port = service.getServiciosFlotaPort();
        return port.cargarArchivo(pathArchivo);
    }
    
}
