# customer-research-tutorial
Automatically research and outbound companies with Exa API and google sheets app scripts. 

## GETTING STARTED
1. Exa API key: get started for free [here](dashboard.exa.ai).
2. Create an OpenAI key
3. Make a copy of this example spreadsheet: https://docs.google.com/spreadsheets/d/1ZsLlbdgFBFwlhtUC-8yQF31jzEA0JfSfwY2Fnhnhhfk/edit?usp=sharing
4. In the spreadsheet, go to extensions -> app scripts -> add the code from this repo into your app scripts
6. To run the automation, press the blue 'Update Sheet' button that will run 'enrichSheet()' inside main.gs

## CUSTOMIZING
- Add your personal Exa API key and OpenAI key 
- Templates: in the 'templates' tab of the google sheet, add your own customer categories, category descriptions, routing, and templates
- Code Categories: inside the assignCategory() function, update the category enums to be your personal customer categories 
