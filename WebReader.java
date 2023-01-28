package test;
import java.net.*;
import java.io.*;
import org.jsoup.*;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class WebReader {

	private Elements elements = null;
	

	
	public void getContentFromURWithJsoup(String url) {
		try {
			Document doc = Jsoup.connect(url).get();
			this.elements = doc.select("tr");
			for(int i=0; i<this.elements.size();i++) {
				//System.out.println(">>> " + this.elements.get(i).text());
				Elements element = this.elements.select("th,td");
				for(int j=0; j<element.size();j++) {
					System.out.println(element.get(j).text());
				}
			}
		} catch (Exception e) {
			System.out.println("Exception in reading URL with Jsoup : "+e.getMessage());
		}  
	}
	
}
