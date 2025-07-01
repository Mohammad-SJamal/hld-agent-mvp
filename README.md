# hld-agent-mvp

This is a google appscript MVP of our google slides agent! The theory is that our application communicates  
using google slides API and gemini to create a copy of a template with information filled in from a  
user-inputted prompt. Gemini reads the JSON representation of a presentation and extracts information on  
what's written and how information is organized to create a starting point for a new presentation!

Right now, the MVP is not perfect. Here's how you get started:


# 1. Go to desired template and click on appscript and copy   paste appscript.js into the code
![alt text](/readme-images/image.png)

![alt text](/readme-images/image-1.png)

# 2. Add slides api under services
![alt text](/readme-images/image-2.png)

# 3. Under project settings and for script properties add an api key for an AI of your choice. Gemini has free API keys. The property must be "ai_key" 
![alt text](/readme-images/image-3.png)

# 4. Run the POC
![alt text](/readme-images/image-4.png)

# 5. Your before/after should look like this:
Before:
![alt text](/readme-images/image-6.png)

After:
![alt text](/readme-images/image-7.png)