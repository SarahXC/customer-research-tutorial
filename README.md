# customer-research-tutorial
Automatically research and outbound companies with Exa API and google sheets app scripts. 

## GETTING STARTED
1. Create an Exa API key: get started for free [here](dashboard.exa.ai).
2. Create an OpenAI key
3. Make a copy of [this](https://docs.google.com/spreadsheets/d/1ZsLlbdgFBFwlhtUC-8yQF31jzEA0JfSfwY2Fnhnhhfk/edit?usp=sharing) example spreadsheet.
<img width="1154" alt="Screenshot 2024-06-14 at 1 30 06 PM" src="https://github.com/SarahXC/customer-research-tutorial/assets/11271849/e5a6dff1-82fa-4ab2-8fd8-41f0ed40ac45">
4. In the spreadsheet, go to extensions -> app scripts -> add the code from main.gs into your app scripts
   <img width="620" alt="Screenshot 2024-06-14 at 1 27 16 PM" src="https://github.com/SarahXC/customer-research-tutorial/assets/11271849/e573e977-ddf2-4ba0-a125-37a21db47f7d">
5. To run the automation, press the blue 'Update Sheet' button that will run 'enrichSheet()' inside main.gs
<img width="620" alt="Screenshot 2024-06-14 at 1 29 41 PM" src="https://github.com/SarahXC/customer-research-tutorial/assets/11271849/96a82598-e440-44f2-b3cb-4a11ed55f403">

## CUSTOMIZING
- Add your personal Exa API key and OpenAI key 
- Templates: in the 'templates' tab of the google sheet, add your own customer categories, category descriptions, routing, and templates
- Code Categories: inside the assignCategory() function, update the category enums to be your personal customer categories 
