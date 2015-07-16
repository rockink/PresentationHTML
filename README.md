# PresentationHTML



This converts the Open XML presentation format (pptx),  to HTML, with just CSS and html involved. It doesn't use javascript. What it does is following:

  1. Converts the picture right in the position it is.
  2. Put the texts as it is.
  3. Lists (ordered List, unordered List)
  4. Color of background (normal, not blended, or background picture)
  5. Text-color
  

It just creates something that looks cool in HTML. Here is how I used. 


         PresentationConvert convert = new PresentationConvert(new FileInputStream("/path/to/file.pptx"));
         StringBuilder sb = new StringBuilder();
         for (int i = 0; i < convert.getSlideCount(); i++) {
            
              sb.append(convert.getSlide(i));
              
         }
         
         String s = sb.toString();
         
         
         
  There are lot to be done, but it feels great though. 
  
  This is what I created, but with some libraries. 
  
  
  
It is developed in Linux, and tested with pptx files using  Libre Office. There are some problems with Microsoft Office Presentation though.
