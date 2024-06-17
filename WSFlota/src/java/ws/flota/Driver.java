package ws.flota;

import java.io.File;
import java.io.FileInputStream;
import java.io.Serializable;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DecimalFormat;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Driver implements Serializable {
    
    private static final long serialVersionUID = 1L;

    private Connection conn;
    private Integer numero_vehiculos;
    
    public Driver() {
        numero_vehiculos = 0;
    }
    
    public Connection openConn() {
        try {
            // InitialContext ctx = new InitialContext();
            // DataSource ds = (DataSource) ctx.lookup("Flota_Jndi");
            // this.conn = ds.getConnection();
            // this.conn = DriverManager.getConnection("jdbc:oracle:thin:@192.200.107.12:1521:payware", "SYSTEM", "oracle01");
			this.conn = DriverManager.getConnection("jdbc:oracle:thin:@192.200.109.50:1521:payware", "SYSTEM", "oracle01");
        } catch (Exception ex) {
            System.out.println("************ ERROR OPEN CONEXCION: " + ex.toString());
        }

        return conn;
    }
    
    public Connection closeConn() {
        try {
            this.conn.close();
        } catch (Exception ex) {
            System.out.println("************ ERROR CLOSE CONEXCION: " + ex.toString());
        }

        return conn;
    }
    
    public void SetAutoCommitConn(Boolean opcion) {
        try {
            this.conn.setAutoCommit(opcion);
        } catch (Exception ex) {
            System.out.println("************ ERROR SETAUTOCOMMIT CONEXCION: " + ex.toString());
        }
    }
    
    public void commitConn() {
        try {
            this.conn.commit();
        } catch(Exception ex) {
            System.out.println("************ ERROR COMMIT CONEXCION: " + ex.toString());
        }
    }
    
    public void rollbackConn() {
        try {
            this.conn.rollback();
        } catch(Exception ex) {
            System.out.println("************ ERROR COMMIT CONEXCION: " + ex.toString());
        }
    }
    
    public void Reciclebin() {
        try {
            Statement stmt = conn.createStatement();
            stmt.executeUpdate("PURGE recyclebin");
            stmt.close();
            
            stmt = conn.createStatement();
            stmt.executeUpdate("ALTER SESSION SET recyclebin = OFF");
            stmt.close();
        } catch(Exception ex) {
            System.out.println("************ ERROR LIMPIANDO LA PAPELERA DE RECICLAJE: " + ex.toString());
        }
    }
    
    public String validar_hojas_excel(String path_archivo) {
        String resultado;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            
            XSSFSheet ws = wb.getSheetAt(1);
            if(ws.getSheetName().trim().equals("Vehiculos")) {
                resultado = "EXITO";
            } else {
                resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
            }
            
            if (resultado.equals("EXITO")) {
                ws = wb.getSheetAt(2);
                if (ws.getSheetName().trim().equals("Vehiculo Tarjetas")) {
                    resultado = "EXITO";
                } else {
                    resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
                }
            }
            
            if (resultado.equals("EXITO")) {
                ws = wb.getSheetAt(3);
                if (ws.getSheetName().trim().equals("Ingreso Estacion")) {
                    resultado = "EXITO";
                } else {
                    resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
                }
            }
            
            if (resultado.equals("EXITO")) {
                ws = wb.getSheetAt(4);
                if (ws.getSheetName().trim().equals("Ingreso Producto")) {
                    resultado = "EXITO";
                } else {
                    resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
                }
            }
            
            if (resultado.equals("EXITO")) {
                ws = wb.getSheetAt(5);
                if (ws.getSheetName().trim().equals("Ingreso Horario")) {
                    resultado = "EXITO";
                } else {
                    resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
                }
            }
            
            if (resultado.equals("EXITO")) {
                ws = wb.getSheetAt(6);
                if (ws.getSheetName().trim().equals("Ingreso Cantidad")) {
                    resultado = "EXITO";
                } else {
                    resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
                }
            }
            
        } catch(Exception ex) {
            resultado = "Error al cargar archivo… El nombre de la hojas dentro del archivo Excel no son correctas.";
        }
        
        return resultado;
    }
    
    // ************************************** VALIDACIONES EN LAS TABLAS SQL ******************************************
    private Boolean existe_registro_estacion(String vehiculo, Integer estacion, Integer branch_i, Integer emisor, Integer producto) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_MERCH_VEHICLE F WHERE F.EMISOR=" + emisor + " AND F.PRODUCTO=" + producto + " AND F.VEHICLE_ID='" + vehiculo + "' AND F.MERCHANT_ID=" + estacion + " AND F.BRANCH_ID=" + branch_i;
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_registro_estacion(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_registro_producto(String vehiculo, Integer emisor, Integer producto) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_VEHICLE_PROD_REST F WHERE F.EMISOR=" + emisor + " AND F.PRODUCTO=" + producto + " AND F.VEHICLE_ID='" + vehiculo + "'";
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_registro_producto(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_registro_horario(String vehiculo, Integer emisor, Integer producto, Integer day) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_VEHICLE_CONTROL F WHERE F.EMISOR=" + emisor + " AND F.PRODUCTO=" + producto + " AND F.VEHICLE_ID='" + vehiculo + "' AND F.DAY=" + day;
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_registro_horario(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_actualizacion_cantidad(String vehiculo, Integer emisor, Integer producto) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_VEHICLE_RESTRICTIONS F WHERE F.EMISOR=" + emisor + " AND F.PRODUCTO=" + producto + " AND F.VEHICLE_ID='" + vehiculo + "'";
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_actualizacion_cantidad(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_vehiculo(String vehiculo, Integer emisor, Integer producto) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_VEHICLE F WHERE F.ISSUER=" + emisor + " AND F.PRODUCT=" + producto + " AND F.VEHICLE_ID='" + vehiculo + "'";
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_vehiculo(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_producto(Integer emisor, Integer producto, Integer product_code) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_PRODUCT_CODES F WHERE F.ISSUER=" + emisor + " AND F.PRODUCT=" + producto + " AND F.PRODUCT_CODE=" + product_code;
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_producto(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_vehiculo_tarjeta(String tarjeta, String vehicle_id, Integer emisor, Integer sucursal_emisor, Integer producto) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_VEHICLE_CARD F WHERE F.EMISOR=" + emisor + " AND F.PRODUCTO=" + producto + " AND F.SUCURSAL_EMISOR=" + sucursal_emisor + " AND F.TARJETA='" + tarjeta + "' AND F.VEHICLE_ID='" +vehicle_id + "'";
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_vehiculo_tarjeta(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_tarjeta(String tarjeta, Integer producto, Integer sucursal_emisor) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.TARJETAS F WHERE F.PRODUCTO=" + producto + " AND F.TARJETA='" + tarjeta + "' AND F.SUCURSAL_EMISOR=" + sucursal_emisor;
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_tarjeta(): " + ex.toString());
        }
        
        return resultado;
    }
    
    private Boolean existe_reglas(Integer emisor, Integer producto, Integer affinity_group) {
        Boolean resultado = false;
        
        try {
            String cadenasql = "SELECT F.* FROM PRODISSR.FLEET_AUTH_RULES F WHERE F.ISSUER=" + emisor + " AND F.PRODUCT=" + producto + " AND F.AFFINITY_GROUP=" + affinity_group;
            Statement stmt = this.conn.createStatement();
            ResultSet rs = stmt.executeQuery(cadenasql);
            while(rs.next()) {
                resultado = true;
            }
            rs.close();
            stmt.close();
        } catch(Exception ex) {
            resultado = false;
            System.out.println("====> existe_registro_regla():" + ex.toString());
        }
        
        return resultado;
    }
    
    // ************************************** CARGA DE INFORMACION EN LAS TABLAS **************************************
    public String cargar_vehiculo(String path_archivo) {
        String resultado = "";
        Integer linea = 0;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(1);
            Integer filas = ws.getLastRowNum() + 1;
            for(Integer i=1; i < filas; i++) {
                linea = i + 1;
                
                XSSFRow row = ws.getRow(i);
                XSSFCell vehicle_id = row.getCell(0);
                XSSFCell issuer = row.getCell(1);
                XSSFCell product = row.getCell(2);
                XSSFCell affinity_group = row.getCell(3);
                XSSFCell register_id = row.getCell(4);
                XSSFCell description = row.getCell(5);
                XSSFCell brand = row.getCell(6);
                XSSFCell year = row.getCell(7);
                XSSFCell kilometers = row.getCell(8);
                
                DecimalFormat formatter = new DecimalFormat("#");
                String x_vehicle_id = "";
                String x_register_id = "";
                String x_description = "";
                String x_brand = "";
                Integer x_issuer = 0;
                Integer x_product = 0;
                Integer x_affinity_group = 0;
                Integer x_year = 0;
                Integer x_kilometers = 0;
                
                
                try {
                    x_vehicle_id = vehicle_id.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo VEHICLE_ID no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_register_id = register_id.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo REGISTER_ID no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_description = description.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo DESCRIPCION no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_brand = brand.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo BRAND no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_issuer = Integer.parseInt(formatter.format(Double.parseDouble(issuer.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo ISSUER debe ser numérico ... linea: " + linea);
                }
                try {
                    x_product = Integer.parseInt(formatter.format(Double.parseDouble(product.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo PRODUCT debe ser numérico ... linea: " + linea);
                }
                try {
                    x_affinity_group = Integer.parseInt(formatter.format(Double.parseDouble(affinity_group.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo AFFINITY_GROUP debe ser numérico ... linea: " + linea);
                }
                try {
                    x_year = Integer.parseInt(formatter.format(Double.parseDouble(year.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo YEAR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_kilometers = Integer.parseInt(formatter.format(Double.parseDouble(kilometers.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo... El valor del campo KILOMETERS debe ser numérico ... linea: " + linea);
                }
                
                //VALIDACIONES DE LA TRANSACCION.
                if(existe_vehiculo(x_vehicle_id,x_issuer,x_product)) {
                    throw new Exception("Carga Vehículo... El vehículo ya existe en Payware ... linea: " + linea);
                }
                
                if(!existe_reglas(x_issuer,x_product,x_affinity_group)) {
                    throw new Exception("Carga Vehículo... No existen las reglas para el país en Payware ... linea: " + linea);
                }
                
                String cadenasql = "INSERT INTO PRODISSR.FLEET_VEHICLE ("
                        + "VEHICLE_ID, "
                        + "ISSUER, "
                        + "PRODUCT, "
                        + "AFFINITY_GROUP, "
                        + "REGISTER_ID, "
                        + "DESCRIPTION, "
                        + "BRAND, "
                        + "YEAR, "
                        + "KILOMETERS) VALUES ('" 
                        + x_vehicle_id + "',"
                        + x_issuer + ","
                        + x_product + ","
                        + x_affinity_group + ",'"
                        + x_register_id + "','"
                        + x_description + "','"
                        + x_brand + "',"
                        + x_year + ","
                        + x_kilometers + ")";
                Statement stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
                
                cadenasql = "INSERT INTO PRODISSR.FLEET_VEHICLE_RESTRICTIONS ("
                        + "VEHICLE_ID, "
                        + "DAILY_QUANTITY, "
                        + "DAILY_AMOUNT, "
                        + "WEEKLY_QUANTITY, "
                        + "WEEKLY_AMOUNT, "
                        + "FORTNIGHTLY_QUANTITY, "
                        + "FORTNIGHTLY_AMOUNT, "
                        + "MONTHLY_QUANTITY, "
                        + "MONTHLY_AMOUNT, "
                        + "MAX_QUANTITY, "
                        + "MAX_AMOUNT, "
                        + "EMISOR, "
                        + "PRODUCTO, "
                        + "MIN_KM) VALUES ('" 
                        + x_vehicle_id + "',"
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + "0" + ","
                        + x_issuer + ","
                        + x_product + ","
                        + "0" + ")";
                stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
            }
            
            this.numero_vehiculos = linea - 1;
            resultado = "EXITO";
        } catch(Exception ex) {
            resultado = ex.toString();
            resultado = resultado.replaceAll("java.lang.Exception: ", "");
        }
        
        return resultado;
    }
    
    public String cargar_vehiculo_tarjeta(String path_archivo) {
        String resultado = "";
        Integer linea = 0;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(2);
            Integer filas = ws.getLastRowNum() + 1;
            for(Integer i=1; i < filas; i++) {
                linea = i + 1;
                
                XSSFRow row = ws.getRow(i);
                XSSFCell emisor = row.getCell(0);
                XSSFCell sucursal_emisor = row.getCell(1);
                XSSFCell producto = row.getCell(2);
                XSSFCell tarjeta = row.getCell(3);
                XSSFCell vehicle_id = row.getCell(4);
                
                DecimalFormat formatter = new DecimalFormat("#");
                String x_tarjeta = "";
                String x_vehicle_id = "";
                Integer x_emisor = 0;
                Integer x_sucursal_emisor = 0;
                Integer x_producto = 0;
                
                
                try {
                    x_tarjeta = tarjeta.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo Tarjeta... El valor del campo TARJETA no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_vehicle_id = vehicle_id.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo Tarjeta... El valor del campo VEHICLE_ID no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_emisor = Integer.parseInt(formatter.format(Double.parseDouble(emisor.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo Tarjeta... El valor del campo EMISOR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_sucursal_emisor = Integer.parseInt(formatter.format(Double.parseDouble(sucursal_emisor.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo Tarjeta... El valor del campo SUCURSAL_EMISOR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_producto = Integer.parseInt(formatter.format(Double.parseDouble(producto.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Vehículo Tarjeta... El valor del campo PRODUCTO debe ser numérico ... linea: " + linea);
                }
                
                //VALIDACIONES DE LA TRANSACCION.
                if(existe_vehiculo_tarjeta(x_tarjeta, x_vehicle_id, x_emisor, x_sucursal_emisor, x_producto)) {
                    throw new Exception("Carga Vehículo Tarjeta... Registro duplicado ... linea: " + linea);
                }
                
                if(!existe_vehiculo(x_vehicle_id,x_emisor,x_producto)) {
                    throw new Exception("Carga Vehículo Tarjeta... El vehículo no existe en Payware ... linea: " + linea);
                }
                
                if(!existe_tarjeta(x_tarjeta, x_producto, x_sucursal_emisor)) {
                    throw new Exception("Carga Vehículo Tarjeta... La tarjeta no existe en Payware ... linea: " + linea);
                }
                
                String cadenasql = "INSERT INTO PRODISSR.FLEET_VEHICLE_CARD ("
                        + "EMISOR, "
                        + "SUCURSAL_EMISOR, "
                        + "PRODUCTO, "
                        + "TARJETA, "
                        + "VEHICLE_ID) VALUES ("
                        + x_emisor + ","
                        + x_sucursal_emisor + ","
                        + x_producto + ",'"
                        + x_tarjeta + "','"
                        + x_vehicle_id + "')";
                Statement stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
            }
            
            resultado = "EXITO";
        } catch(Exception ex) {
            resultado = ex.toString();
            resultado = resultado.replaceAll("java.lang.Exception: ", "");
        }
        
        return resultado;
    }
    
    public String cargar_estacion(String path_archivo) {
        String resultado = "";
        Integer linea = 0;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(3);
            Integer filas = ws.getLastRowNum() + 1;
            for(Integer i=1; i < filas; i++) {
                linea = i + 1;
                
                XSSFRow row = ws.getRow(i);
                XSSFCell vehiculo = row.getCell(0);
                XSSFCell estacion = row.getCell(1);
                XSSFCell branch_i = row.getCell(2);
                XSSFCell emisor = row.getCell(3);
                XSSFCell producto = row.getCell(4);
                
                DecimalFormat formatter = new DecimalFormat("#");
                String x_vehiculo = "";
                Integer x_estacion = 0;
                Integer x_branch_i = 0;
                Integer x_emisor = 0;
                Integer x_producto = 0;
                
                
                try {
                    x_vehiculo = vehiculo.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Estación... El valor del campo VEHICULO no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_estacion = Integer.parseInt(formatter.format(Double.parseDouble(estacion.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Estación... El valor del campo ESTACION debe ser numérico ... linea: " + linea);
                }
                try {
                    x_branch_i = Integer.parseInt(formatter.format(Double.parseDouble(branch_i.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Estación... El valor del campo BRANCH_ID debe ser numérico ... linea: " + linea);
                }
                try {
                    x_emisor = Integer.parseInt(formatter.format(Double.parseDouble(emisor.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Estación... El valor del campo EMISOR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_producto = Integer.parseInt(formatter.format(Double.parseDouble(producto.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Estación... El valor del campo PRODUCTO debe ser numérico ... linea: " + linea);
                }
                
                //VALIDACIONES DE LA TRANSACCION.
                if(existe_registro_estacion(x_vehiculo, x_estacion, x_branch_i, x_emisor, x_producto)) {
                    throw new Exception("Carga Estación... Registro duplicado ... linea: " + linea);
                }
                
                if(!existe_vehiculo(x_vehiculo,x_emisor,x_producto)) {
                    throw new Exception("Carga Estación... El vehículo no existe en Payware ... linea: " + linea);
                }
                
                if(x_branch_i!=1) {
                    throw new Exception("Carga Estación... El valor de campo BRANCH_ID debe ser 1 ... linea: " + linea);
                }
                
                String cadenasql = "INSERT INTO PRODISSR.FLEET_MERCH_VEHICLE ("
                        + "VEHICLE_ID,"
                        + "MERCHANT_ID,"
                        + "BRANCH_ID,"
                        + "EMISOR,"
                        + "PRODUCTO) VALUES ('" 
                        + x_vehiculo + "',"
                        + x_estacion + ","
                        + x_branch_i + ","
                        + x_emisor + ","
                        + x_producto + ")";
                Statement stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
            }
            
            resultado = "EXITO";
        } catch(Exception ex) {
            resultado = ex.toString();
            resultado = resultado.replaceAll("java.lang.Exception: ", "");
        }
        
        return resultado;
    }
    
    public String cargar_producto(String path_archivo) {
        String resultado = "";
        Integer linea = 0;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(4);
            Integer filas = ws.getLastRowNum() + 1;
            for(Integer i=1; i < filas; i++) {
                linea = i + 1;
                
                XSSFRow row = ws.getRow(i);
                XSSFCell vehiculo = row.getCell(0);
                XSSFCell producto_code = row.getCell(1);
                XSSFCell emisor = row.getCell(2);
                XSSFCell producto = row.getCell(3);
                
                DecimalFormat formatter = new DecimalFormat("#");
                String x_vehiculo = "";
                Integer x_producto_code = 0;
                Integer x_emisor = 0;
                Integer x_producto = 0;
                
                
                try {
                    x_vehiculo = vehiculo.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Producto... El valor del campo VEHICULO no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_producto_code = Integer.parseInt(formatter.format(Double.parseDouble(producto_code.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Producto... El valor del campo PRODUCT_CODE debe ser numérico ... linea: " + linea);
                }
                try {
                    x_emisor = Integer.parseInt(formatter.format(Double.parseDouble(emisor.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Producto... El valor del campo EMISOR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_producto = Integer.parseInt(formatter.format(Double.parseDouble(producto.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Producto... El valor del campo PRODUCTO debe ser numérico ... linea: " + linea);
                }
                
                //VALIDACIONES DE LA TRANSACCION.
                if(existe_registro_producto(x_vehiculo,x_emisor,x_producto)) {
                    throw new Exception("Carga Producto... Registro duplicado ... linea: " + linea);
                }
                
                if(!existe_vehiculo(x_vehiculo,x_emisor,x_producto)) {
                    throw new Exception("Carga Producto... El vehículo no existe en Payware ... linea: " + linea);
                }
                
                if(!existe_producto(x_emisor,x_producto,x_producto_code)) {
                    throw new Exception("Carga Producto... El producto no existe en Payware ... linea: " + linea);
                }
                
                String cadenasql = "INSERT INTO PRODISSR.FLEET_VEHICLE_PROD_REST ("
                        + "VEHICLE_ID,"
                        + "PRODUCT_CODE,"
                        + "EMISOR,"
                        + "PRODUCTO) VALUES ('" 
                        + x_vehiculo + "',"
                        + x_producto_code + ","
                        + x_emisor + ","
                        + x_producto + ")";
                Statement stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
            }
            
            resultado = "EXITO";
        } catch(Exception ex) {
            resultado = ex.toString();
            resultado = resultado.replaceAll("java.lang.Exception: ", "");
        }
        
        return resultado;
    }
    
    public String cargar_horario(String path_archivo) {
        String resultado = "";
        Integer linea = 0;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(5);
            Integer filas = ws.getLastRowNum() + 1;
            for(Integer i=1; i < filas; i++) {
                linea = i + 1;
                
                XSSFRow row = ws.getRow(i);
                XSSFCell vehiculo = row.getCell(0);
                XSSFCell day = row.getCell(1);
                XSSFCell hour_from = row.getCell(2);
                XSSFCell hour_to = row.getCell(3);
                XSSFCell emisor = row.getCell(4);
                XSSFCell producto = row.getCell(5);
                
                DecimalFormat formatter = new DecimalFormat("#");
                String x_vehiculo = "";
                Integer x_day = 0;
                Integer x_hour_from = 0;
                Integer x_hour_to = 0;
                Integer x_emisor = 0;
                Integer x_producto = 0;
                
                
                try {
                    x_vehiculo = vehiculo.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Carga Horario... El valor del campo VEHICULO no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_day = Integer.parseInt(formatter.format(Double.parseDouble(day.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Horario... El valor del campo DAY debe ser numérico ... linea: " + linea);
                }
                try {
                    x_hour_from = Integer.parseInt(formatter.format(Double.parseDouble(hour_from.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Horario... El valor del campo HOUR_FROM debe ser numérico ... linea: " + linea);
                }
                try {
                    x_hour_to = Integer.parseInt(formatter.format(Double.parseDouble(hour_to.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Horario... El valor del campo HOUR_TO debe ser numérico ... linea: " + linea);
                }
                try {
                    x_emisor = Integer.parseInt(formatter.format(Double.parseDouble(emisor.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Horario... El valor del campo EMISOR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_producto = Integer.parseInt(formatter.format(Double.parseDouble(producto.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Carga Horario... El valor del campo PRODUCTO debe ser numérico ... linea: " + linea);
                }
                
                //VALIDACIONES DE LA TRANSACCION.
                if(existe_registro_horario(x_vehiculo,x_emisor,x_producto, x_day)) {
                    throw new Exception("Carga Horario... Registro duplicado ... linea: " + linea);
                }
                
                if(!existe_vehiculo(x_vehiculo,x_emisor,x_producto)) {
                    throw new Exception("Carga Horario... El vehículo no existe en Payware ... linea: " + linea);
                }
                
                String cadenasql = "INSERT INTO PRODISSR.FLEET_VEHICLE_CONTROL ("
                        + "VEHICLE_ID,"
                        + "DAY,"
                        + "HOUR_FROM,"
                        + "HOUR_TO,"
                        + "EMISOR,"
                        + "PRODUCTO) VALUES ('" 
                        + x_vehiculo + "',"
                        + x_day + ","
                        + x_hour_from + ","
                        + x_hour_to + ","
                        + x_emisor + ","
                        + x_producto + ")";
                Statement stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
            }
            
            resultado = "EXITO";
        } catch(Exception ex) {
            resultado = ex.toString();
            resultado = resultado.replaceAll("java.lang.Exception: ", "");
        }
        
        return resultado;
    }
    
    public String actualiza_cantidad(String path_archivo) {
        String resultado = "";
        Integer linea = 0;
        
        try {
            File excel = new File(path_archivo);
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(6);
            Integer filas = ws.getLastRowNum() + 1;
            for(Integer i=1; i < filas; i++) {
                linea = i + 1;
                
                XSSFRow row = ws.getRow(i);
                XSSFCell vehiculo = row.getCell(0);
                XSSFCell daily_quantity = row.getCell(1);
                XSSFCell daily_amount = row.getCell(2);
                XSSFCell weekly_quantity = row.getCell(3);
                XSSFCell weekly_amount = row.getCell(4);
                XSSFCell fortnightly_quantity = row.getCell(5);
                XSSFCell fortnightly_amount = row.getCell(6);
                XSSFCell monthly_quantity = row.getCell(7);
                XSSFCell monthly_amount = row.getCell(8);
                XSSFCell max_quantity = row.getCell(9);
                XSSFCell max_amount = row.getCell(10);
                XSSFCell emisor = row.getCell(11);
                XSSFCell producto = row.getCell(12);
                XSSFCell min_km = row.getCell(13);

                DecimalFormat formatter = new DecimalFormat("#");
                DecimalFormat formatoDouble = new DecimalFormat("#0.0000");
                String x_vehiculo = "";
                Double x_daily_quantity = 0.00;
                Double x_daily_amount = 0.00;
                Double x_weekly_quantity = 0.00;
                Double x_weekly_amount = 0.00;
                Double x_fortnightly_quantity = 0.00;
                Double x_fortnightly_amount = 0.00;
                Double x_monthly_quantity = 0.00;
                Double x_monthly_amount = 0.00;
                Double x_max_quantity = 0.00;
                Double x_max_amount = 0.00;
                Integer x_emisor = 0;
                Integer x_producto = 0;
                Integer x_min_km = 0;
                
                
                try {
                    x_vehiculo = vehiculo.toString().trim();
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo VEHICULO no puede ser nulo ... linea: " + linea);
                }
                try {
                    x_daily_quantity = Double.parseDouble(formatoDouble.format(Double.parseDouble(daily_quantity.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo DAILY_QUANTITY debe ser numérico ... linea: " + linea);
                }
                try {
                    x_daily_amount = Double.parseDouble(formatoDouble.format(Double.parseDouble(daily_amount.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo DAILY_AMOUNT debe ser numérico ... linea: " + linea);
                }
                try {
                    x_weekly_quantity = Double.parseDouble(formatoDouble.format(Double.parseDouble(weekly_quantity.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo WEEKLY_QUANTITY debe ser numérico ... linea: " + linea);
                }
                try {
                    x_weekly_amount = Double.parseDouble(formatoDouble.format(Double.parseDouble(weekly_amount.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo WEEKLY_AMOUNT debe ser numérico ... linea: " + linea);
                }
                try {
                    x_fortnightly_quantity = Double.parseDouble(formatoDouble.format(Double.parseDouble(fortnightly_quantity.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo FORTNIGHTLY_QUANTITY debe ser numérico ... linea: " + linea);
                }
                try {
                    x_fortnightly_amount = Double.parseDouble(formatoDouble.format(Double.parseDouble(fortnightly_amount.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo FORTNIGHTLY_AMOUNT debe ser numérico ... linea: " + linea);
                }
                try {
                    x_monthly_quantity = Double.parseDouble(formatoDouble.format(Double.parseDouble(monthly_quantity.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo MONTHLY_QUANTITY debe ser numérico ... linea: " + linea);
                }
                try {
                    x_monthly_amount = Double.parseDouble(formatoDouble.format(Double.parseDouble(monthly_amount.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo MONTHLY_AMOUNT debe ser numérico ... linea: " + linea);
                }
                try {
                    x_max_quantity = Double.parseDouble(formatoDouble.format(Double.parseDouble(max_quantity.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo MAX_QUANTITY debe ser numérico ... linea: " + linea);
                }
                try {
                    x_max_amount = Double.parseDouble(formatoDouble.format(Double.parseDouble(max_amount.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo MAX_AMOUNT debe ser numérico ... linea: " + linea);
                }
                try {
                    x_emisor = Integer.parseInt(formatter.format(Double.parseDouble(emisor.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo EMISOR debe ser numérico ... linea: " + linea);
                }
                try {
                    x_producto = Integer.parseInt(formatter.format(Double.parseDouble(producto.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo PRODUCTO debe ser numérico ... linea: " + linea);
                }
                try {
                    x_min_km = Integer.parseInt(formatter.format(Double.parseDouble(min_km.toString().trim())));
                } catch(Exception ex) {
                    throw new Exception("Actualización Cantidad... El valor del campo MIN_KM debe ser numérico ... linea: " + linea);
                }
                
                //VALIDACIONES DE LA TRANSACCION.
                if(!existe_actualizacion_cantidad(x_vehiculo,x_emisor,x_producto)) {
                    throw new Exception("Actualización Cantidad... El registro (VEHICULO,EMISOR,PRODUCTO) no existe ... linea: " + linea);
                }
                
                String cadenasql = "UPDATE PRODISSR.FLEET_VEHICLE_RESTRICTIONS SET "
                        + "DAILY_QUANTITY=" + x_daily_quantity + ", "
                        + "DAILY_AMOUNT=" + x_daily_amount + ", "
                        + "WEEKLY_QUANTITY=" + x_weekly_quantity + ", "
                        + "WEEKLY_AMOUNT=" + x_weekly_amount + ", "
                        + "FORTNIGHTLY_QUANTITY=" + x_fortnightly_quantity + ", "
                        + "FORTNIGHTLY_AMOUNT=" + x_fortnightly_amount + ", "
                        + "MONTHLY_QUANTITY=" + x_monthly_quantity + ", "
                        + "MONTHLY_AMOUNT=" + x_monthly_amount + ", "
                        + "MAX_QUANTITY=" + x_max_quantity + ", "
                        + "MAX_AMOUNT=" + x_max_amount + ", "
                        + "MIN_KM=" + x_min_km + " "
                        + "WHERE VEHICLE_ID='" + x_vehiculo +"' AND EMISOR=" + x_emisor +" AND PRODUCTO=" + x_producto;
                Statement stmt = this.conn.createStatement();
                stmt.executeUpdate(cadenasql);
                stmt.close();
            }
            
            resultado = "EXITO";
        } catch(Exception ex) {
            resultado = ex.toString();
            resultado = resultado.replaceAll("java.lang.Exception: ", "");
        }
        
        return resultado;
    }

    // ************************************** GETTERS & SETTERS *******************************************************
    public Integer getNumero_vehiculos() {
        return numero_vehiculos;
    }

    public void setNumero_vehiculos(Integer numero_vehiculos) {
        this.numero_vehiculos = numero_vehiculos;
    }
    
}
