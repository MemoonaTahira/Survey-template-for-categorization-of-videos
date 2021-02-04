# Survey-template-for-categorization-of-videos
This is a small VBA based data labelling application for categorizing videos/iamge/text


- Overview: 

     - A powerpoint slide that ask observer to enter his name to identify him/her later.
     - Next slide that guides the observer what to do with a set of instructions. 
     - A start button to start with the labelling
     - With one video on each slide with 3 possible lables, one of which the observer has to pick. By default, observer has to pick or he won't be allowed to move ahead. 
     - An end button that will save the updated excel file with the new labels this observer has picked

- How to use: 

Step 1: Edit your labels by going to Developer tools in the ribbon and editing the code.

Step 2: Add your video to the slides. You can add image or text or whatever else you want labelled as well.

Step 3: Make sure your excel file to store the results is on the path (i.e. in the same folder as the slides) and closed before running the slideshow. The VBA code will automatically call it up.

Step 4: Start the slideshow. Use forward arrow to move forward once observer has picked his choice.

Step 5: If the observer exits powerpoint in the middle of the labelling, close the excel sheet before resatrting the process. His/Her partial labels will still be retained in the excel sheet.



- Additional: Screenshot of a sample slide. This was developed to categorize the videos of [CrowdFix](https://github.com/MemoonaTahira/CrowdFix) dataset into three categories: Sparse, Dense and Free Flowing, Dense Congested.
