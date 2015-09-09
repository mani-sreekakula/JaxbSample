package com.mani.app;

import java.io.StringWriter;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;

public class CustomParser {

	public static String marshall(Object object) {
		String result = "";
		JAXBContext jaxbContext;
		try {
			jaxbContext = JAXBContext.newInstance(new Class[] { object.getClass() });
			StringWriter writer = new StringWriter();
			Marshaller jaxbMarshaller = jaxbContext.createMarshaller();
			jaxbMarshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
			jaxbMarshaller.marshal(object, writer);
			result = writer.toString();
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		return result;
	}
}
