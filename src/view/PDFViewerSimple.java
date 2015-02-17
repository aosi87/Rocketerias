/*
 * Copyright 2015 Elpapo.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package view;

// PDFViewer.java

import java.awt.*;
import java.awt.geom.*;

import java.io.*;

import java.nio.*;
import java.nio.channels.*;

import javax.swing.*;

import com.sun.pdfview.*;

public class PDFViewerSimple extends JFrame
{
   static Image image;

   public PDFViewerSimple (String title){
      super (title);
      setDefaultCloseOperation (JDialog.DISPOSE_ON_CLOSE);

      JLabel label = new JLabel (new ImageIcon (image));
      label.setVerticalAlignment (JLabel.TOP);

      setContentPane (new JScrollPane (label));

      pack ();
      setVisible (true);
   }

   public static void init (String file, int pageNum) throws IOException   {
      //if (args.length < 1 || args.length > 2)
      //{
        //  System.err.println ("usage: java PDFViewer pdfspec [pagenum]");
          //return;
      //}
      final String filePath = file;
      int pagenum = pageNum;//(args.length == 1) ? 1 : Integer.parseInt (args [1]);
      if (pagenum < 1)
          pagenum = 1;

      RandomAccessFile raf = new RandomAccessFile (new File (file),"r");//args [0]), "r");
      FileChannel fc = raf.getChannel ();
      ByteBuffer buf = fc.map (FileChannel.MapMode.READ_ONLY, 0, fc.size ());
      PDFFile pdfFile = new PDFFile (buf);

      int numpages = pdfFile.getNumPages ();
      System.out.println ("Number of pages = "+numpages);
      if (pagenum > numpages)
          pagenum = numpages;

      PDFPage page = pdfFile.getPage (pagenum);
              
      Rectangle2D r2d = page.getBBox ();

      double width = r2d.getWidth ();
      double height = r2d.getHeight ();
      width /= 72.0;
      height /= 72.0;
      int res = Toolkit.getDefaultToolkit ().getScreenResolution ();
      width *= res;
      height *= res;

      image = page.getImage ((int) width, (int) height, r2d, null, true, true);

      Runnable r = new Runnable ()
                   {
                       public void run ()
                       {
                          new PDFViewerSimple ("PDF Viewer: "+ filePath);//args [0]);
                       }
                   };
      EventQueue.invokeLater (r);
   }
}
