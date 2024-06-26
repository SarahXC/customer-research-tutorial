# customer-research-tutorial
Automatically research and outbound companies with the Exa API and google sheets app scripts. Watch the tutorial [here](https://www.loom.com/share/69bbc99a7a43458490b88bbd21b945a1?sid=c7eb94e3-4be1-4fc0-a201-b0672cf7a43c).

## GETTING STARTED
1. Create an Exa API key: get started for free [here](https://dashboard.exa.ai).
2. Create an OpenAI key
3. Make a copy of [this](https://docs.google.com/spreadsheets/d/1ZsLlbdgFBFwlhtUC-8yQF31jzEA0JfSfwY2Fnhnhhfk/edit?usp=sharing) example spreadsheet.

4. In the spreadsheet, go to extensions -> app scripts -> add the code from main.gs into your app scripts
   <img width="620" alt="Screenshot 2024-06-14 at 1 27 16 PM" src="https://github.com/SarahXC/customer-research-tutorial/assets/11271849/e573e977-ddf2-4ba0-a125-37a21db47f7d">
5. Add your personal Exa API key and OpenAI key 
5. To run the automation, press the blue 'Update Sheet' button that will run 'enrichSheet()' inside main.gs
   <img width="620" alt="Screenshot 2024-06-14 at 1 29 41 PM" src="https://github.com/SarahXC/customer-research-tutorial/assets/11271849/96a82598-e440-44f2-b3cb-4a11ed55f403">

## CUSTOMIZING
- Templates: in the 'templates' tab of the google sheet, add your own customer categories, category descriptions, routing, and templates
- Code categories: inside the assignCategory() function, update the category enums to be your personal customer categories 
