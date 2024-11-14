# Step 0
Ask GPT-4o "i am building gpt extension for google docs what is the best package to use?"
see reply in `0. draft PRD.md` 

# Step 1
create PRD in `instruction.md` file using the following structure
```
## Project Overview
xxxx

## Core functionalities
xxxx

## Documentation
xxxx

## Project file structure

```

# Step 2 (optional)
Enhance PRD in `instruction.md` using `gpt o1-preview` model
Did not use this optional step here

# Composer (using new `claude-3-5-sonnet-20241022` model)
Основная идея - создаем goole doc extension по шагам и на каждом шаге проверяем как работает, если какая проблема возникает - траблшутим про помощи Composer, и только после устранения ошибки на текущем шаге переходим на следующий. Последовательность шагов определена в PRD.

##  Step 1
start creating gpt extension for google docs, let's build ## Connecting Google Apps Script to OpenAI API

## Step 2
great, let's continue and build ## Setting Up the Sidebar UI (ui.html)

## Step 3
Все круто! Переведи пожалуйста все опции меню, которые уместно, на русский язык, не меняй логику работы!

## Step 4
when I click "Settings" I get an error "Exception: HTML-файл с именем settings не найден." I need to able to configure the following parameters:
- base url of openai compatible model
- model itself (either chose from the short list of 4 most popular openai models or enter manually the name of the model)
- temperature (from 0 to 1, of possible use progress bar or something like this)
- max tokens (from 150 till infinity)

## Step 5
Please create a comprehensive  github README.md file in Russian based on PRD provided. 