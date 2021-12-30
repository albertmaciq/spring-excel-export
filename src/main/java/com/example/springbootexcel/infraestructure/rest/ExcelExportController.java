package com.example.springbootexcel.infraestructure.rest;

import org.springframework.web.bind.annotation.GetMapping;

public class ExcelExportController {

  @GetMapping("/")
  public String excel() {
    return "Hello export-excel";
  }
}
