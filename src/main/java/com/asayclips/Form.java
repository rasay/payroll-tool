package com.asayclips;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.HPos;
import javafx.scene.Scene;
import javafx.scene.control.*;

import javafx.scene.layout.Background;
import javafx.scene.layout.GridPane;
import javafx.stage.Stage;

public class Form extends Application {
    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Sport Clips Payroll");

        GridPane gridPane = new GridPane();
        gridPane.setHgap(10);
        gridPane.setVgap(10);

        // first row
        final Label datePickerLabel = new Label("Payroll End Date");
        gridPane.add(datePickerLabel, 0, 0);

        final DatePicker datePicker = new DatePicker();
        gridPane.add(datePicker, 0, 1);

        // second row
        final Label storePickerLabel = new Label("Store #");
        gridPane.add(storePickerLabel, 0, 3);

        final TextField storeText = new TextField();
        gridPane.add(storeText, 0, 4);

        // message area
        final TextArea messageArea = new TextArea();
        messageArea.backgroundProperty().setValue(Background.EMPTY);
        gridPane.add(messageArea, 0, 6);

        //bottom
        Button goButton = new Button();
        goButton.setText("Process");
        goButton.setOnAction(new EventHandler<ActionEvent>() {

            public void handle(ActionEvent event) {
                App app = new App(messageArea);
                app.generatePayroll(datePicker.getEditor().getText(), storeText.getText());
            }
        });
        gridPane.add(goButton, 0, 15);


        Button exitButton = new Button();
        exitButton.setText("Exit");
        exitButton.setOnAction(new EventHandler<ActionEvent>() {

            public void handle(ActionEvent event) {
                System.exit(0);
            }
        });
        gridPane.add(exitButton, 0, 16);

        primaryStage.setScene(new Scene(gridPane, 600, 500));
        primaryStage.show();
    }
}