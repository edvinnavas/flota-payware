package com.flota;

import java.io.File;
import java.io.Serializable;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import javax.annotation.PostConstruct;
import javax.faces.application.FacesMessage;
import javax.faces.bean.ManagedBean;
import javax.faces.bean.ViewScoped;
import javax.faces.context.FacesContext;
import org.apache.commons.io.FileUtils;
import org.primefaces.model.UploadedFile;

@ManagedBean(name = "Index")
@ViewScoped
public class Index implements Serializable {
    
    private static final long serialVersionUID = 1L;

    private UploadedFile file;

    @PostConstruct
    public void init() {
        this.file = null;
    }

    public void cargar_archivo() {
        try {
            Long tamano = this.file.getSize();
            if (tamano > 0) {
                String nombre_archivo = this.file.getFileName().trim();
                Integer tamano_nombre_archivo = nombre_archivo.length();
                if (tamano_nombre_archivo >= 5) {
                    String extension_archivo = nombre_archivo.substring(nombre_archivo.length() - 5, nombre_archivo.length());
                    if (extension_archivo.trim().toLowerCase().equals(".xlsx")) {
                        Driver driver = new Driver();
                        Calendar cal = Calendar.getInstance();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
                        File destFile = new File(driver.getPath_rep_pentaho() + "flota_" + sdf.format(cal.getTime()) + ".xlsx");
                        FileUtils.copyInputStreamToFile(this.file.getInputstream(), destFile);

                        Servicio servicio = new Servicio();
                        String resultado = servicio.cargarArchivo(driver.getPath_rep_pentaho() + "flota_" + sdf.format(cal.getTime()) + ".xlsx");
                        
                        // File archivo = new File(driver.getPath_rep_pentaho() + "flota_" + sdf.format(cal.getTime()) + ".xlsx");
                        // archivo.delete();
                        
                        if(resultado.trim().substring(0,5).equals("EXITO")) {
                            String numero_vehiculos = resultado.trim().substring(6,resultado.trim().length());
                            FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_INFO, "Mensaje del sistema...", "Información cargada correctamente, RESUMEN VEHÍCULOS INSERTADOS: " + numero_vehiculos));
                        } else {
                            FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_ERROR, "Mensaje del sistema...", resultado));
                        }
                    } else {
                        FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_ERROR, "Validación archivo excel...", "Error archivo no ejecutado….. tipo de archivo no permitido (.xlsx)."));
                    }
                } else {
                    FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_ERROR, "Validación archivo excel...", "Error archivo no ejecutado….. tipo de archivo no permitido (.xlsx)."));
                }
            } else {
                FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_ERROR, "Validación archivo excel...", "Error archivo no ejecutado….. tipo de archivo no permitido (.xlsx)."));
            }
        } catch (Exception ex) {
            FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_ERROR, "Mensaje del sistema...", ex.toString()));
        }
    }

    public UploadedFile getFile() {
        return file;
    }

    public void setFile(UploadedFile file) {
        this.file = file;
    }

}
