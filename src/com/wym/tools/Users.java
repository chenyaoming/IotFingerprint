package com.wym.tools;


public class Users {
    private int id;
    private String username;
    private String password;

    public Users(){

    }

    public Users(int id, String username, String password, String UEmail) {
        this.id = id;
        this.username = username;
        this.password = password;
        this.UEmail = UEmail;
    }

    public String getUEmail() {
        return UEmail;
    }

    public void setUEmail(String UEmail) {
        this.UEmail = UEmail;
    }

    private String UEmail;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }



}
