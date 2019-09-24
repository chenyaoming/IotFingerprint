package com.wym.tools;


import java.util.ArrayList;
import java.util.List;

public class UsersDAO {
    public void saveOrUpdate2(Users users) {
    }

    public List findAll() {
        List<Users> usersList = new ArrayList<>();
        usersList.add(new Users(1,"aa","12345","aa@qq.com"));
        usersList.add(new Users(1,"aa","12345","aa@qq.com"));
        usersList.add(new Users(1,"aa","12345","aa@qq.com"));
        return usersList;
    }

    public void save(Users user) {
    }

    public Users findById(int parseInt) {
        return null;
    }
}
