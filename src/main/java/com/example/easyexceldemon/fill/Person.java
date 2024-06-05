package com.example.easyexceldemon.fill;

import lombok.Data;

@Data
public class Person {
    private String name;
    private int age;

    public Person(String name,int age) {
        this.age =age;
        this.name =name;
    }
}

