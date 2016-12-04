package com.malex;

import com.malex.model.DataModel;
import com.malex.service.ExcelWorker;

import java.util.ArrayList;
import java.util.List;

/**
 * Class {link MainApp} run app
 * @author malex
 */
public class MainApp {

    public static void main(String[] args) {

        DataModel dataModel_01 = new DataModel("Max", "St", "Kh", 120.33);
        DataModel dataModel_02 = new DataModel("Alex", "MAx", "Ch", 320.66);
        DataModel dataModel_03 = new DataModel("Anna", "Kov", "Vol", 620.90);

        // fill the list
        List<DataModel> dataModelList = new ArrayList<>();
        dataModelList.add(dataModel_01);
        dataModelList.add(dataModel_02);
        dataModelList.add(dataModel_03);

        //Creating the excel file
        ExcelWorker.createExcelFile("data", "employee",dataModelList);

    }

}
