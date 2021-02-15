package com.sabu.excelapachepoi;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;

@Data
@Getter
@Setter
public class Book {

    private String title;
    private String author;
    private float price;

    public Book() {
    }

    public Book(String title, String author, double price) {
        this.title = title;
        this.author = author;
        this.price = (float) price;
    }

}
