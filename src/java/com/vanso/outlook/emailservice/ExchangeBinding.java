/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.vanso.outlook.emailservice;

import java.net.URI;
import java.net.URISyntaxException;
import java.text.SimpleDateFormat;
import java.util.Date;
import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.MessageBody;
import microsoft.exchange.webservices.data.Recurrence;
import microsoft.exchange.webservices.data.WebCredentials;

/**
 *
 * @author Idris
 */
public class ExchangeBinding {

    public static void main(String arg[]) throws Exception {
        ExchangeBinding xb = new ExchangeBinding();
        ExchangeService service = xb.getExchangeService();
//        EmailMessage msg = new EmailMessage(service);
//        msg.setSubject("Hello world!");
//        msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Managed API."));
//        msg.getToRecipients().add("gentletom2004@gmail.com");
//        msg.send();
//        System.out.println("Message response : " + msg.getIsResend());

        Appointment appointment = new Appointment(service);
        appointment.setSubject("Exchange TEST using Java API - Calendar");
        appointment.setBody(MessageBody.getMessageBodyFromText("Test Body Msg"));

        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date startDate = formatter.parse("2014-01-06 11:01:00");
        Date endDate = formatter.parse("2014-01-06 11:10:00");

        appointment.setStart(startDate);//new Date(2010-1900,5-1,20,20,00));
        appointment.setEnd(endDate); //new Date(2010-1900,5-1,20,21,00));
        appointment.getRequiredAttendees().add("oche.omale@gmail.com");
        appointment.getRequiredAttendees().add("oche.omale@yahoo.co.uk");
        appointment.getRequiredAttendees().add("oche.omale@zenithbank.com");
        appointment.getRequiredAttendees().add("oche.omale@zenithcustodian.com");

//        formatter = new SimpleDateFormat("yyyy-MM-dd");
//        Date recurrenceEndDate = formatter.parse("2010-07-20");
//
//        appointment.setRecurrence(new Recurrence.DailyPattern(appointment.getStart(), 3));
//
//        appointment.getRecurrence().setStartDate(appointment.getStart());
//        appointment.getRecurrence().setEndDate(recurrenceEndDate);
        appointment.save();

    }

    public ExchangeService getExchangeService() {
        // Create the binding.
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials ec = new WebCredentials("Your email address", "Your password");
        try {
            service.setCredentials(ec);
            service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx"));
//            service.autodiscoverUrl("thomas.owoicho@vanso.com");
        } catch (URISyntaxException e) {
            System.out.println("Exception occured : " + e.getMessage());
        }
        return service;
    }
}
