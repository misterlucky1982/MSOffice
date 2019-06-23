package by.misterlucky.msoffice;

import java.io.IOException;

public class ExcelException extends IOException{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public ExcelException(){
		
	}
	
	public ExcelException(String message){
		super(message);
	}
}
