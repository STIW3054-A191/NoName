/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mavenproject2.displayinexcel;

/**
 *
 * @author Lenovo
 */
public class Format {
    private String header;
    private String data;
    
    public Format(){
        
    }
    
    public Format(String header, String data){
        this.header = header;
        this.data = data;
    }

    Format(String column1, String column2, String column3) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
    
    public void setHeader(String header) {
        this.header = header;
    }

    public String getHeader() {
        return header;
    }
    
    public void setData(String data) {
        this.data = data;
    }

    public String getData() {
        return data;
    }
}
