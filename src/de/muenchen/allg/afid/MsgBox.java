/*
* Eine simple MsgBox für UNO-Komponenten, die ohne ServiceManager auskommen müssen.
* Dateiname: MsgBox.java
* Projekt  : n/a
* Funktion : Eine simple MsgBox für UNO-Komponenten.
* 
* Copyright: Landeshauptstadt München
*
* Änderungshistorie:
* Datum      | Wer | Änderungsgrund
* -------------------------------------------------------------------
* 13.09.2005 | LUT | Erstellungpackage de.muenchen.allg.afid;
* ------------------------------------------------------------------- 
*
* @author D-III-ITD 5.1 Christohp Lutz
* @version 1.0
* 
* */
package de.muenchen.allg.afid;

import java.awt.Button;
import java.awt.Dialog;
import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.TextArea;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class MsgBox {

	private static Dialog d;

	public static void simple(String title, String text) {
		Frame window = new Frame();

		d = new Dialog(window, title, true);
		d.setLayout(new FlowLayout(FlowLayout.LEFT));

		Button ok = new Button("OK");
		ok.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// Hide dialog
				MsgBox.d.setVisible(false);
			}
		});

		d.add(new TextArea(text, 20, 80));
		d.add(ok);

		d.pack();
		d.setVisible(true);
	}
}
