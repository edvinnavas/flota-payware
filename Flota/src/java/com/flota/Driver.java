package com.flota;

import java.io.Serializable;

public class Driver implements Serializable {
    
    private static final long serialVersionUID = 1L;
    
    private String path_rep_pentaho;

    public Driver() {
        // path_rep_pentaho = "C:\\\\rep_flota\\";
        path_rep_pentaho = "/root/rep_flota/";
    }

    public String getPath_rep_pentaho() {
        return path_rep_pentaho;
    }

    public void setPath_rep_pentaho(String path_rep_pentaho) {
        this.path_rep_pentaho = path_rep_pentaho;
    }
    
}
