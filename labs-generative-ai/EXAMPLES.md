# `LABS.GENERATIVEAI` samples

Check out the following examples for prompt ideas and to learn how to structure `LABS.GENERATIVEAI` formulas.

1. Analyze public information. Ask a model to summarize a complicated topic.

    ```
    =LABS.GENERATIVEAI("Explain generative AI in a sentence.”)
    ```

1. Process data and define the format. Ask a model to return a particular type of publicly available data and tell it how to format the information. Set the `temperature` value to 0 directly in the function to return consistent responses.

    ```
    =LABS.GENERATIVEAI("Convert these airport codes into a comma-separated list of cities. The codes: LAX, HND, LHR, JNB.", 0)
    ```

    Keep in mind that generative AI models can not guarantee the accuracy of data returned. They are very accurate at creative or language-based tasks, for example, classifying feedback as "positive" or "negative". However, generative AI models have lower accuracy for numerical tasks, such as fetching the population of a city. They have even lower accuracy for tasks requiring numerical comparisons or operations, such as ranking or sorting the results it returns.

1. Use generative AI to reference and analyze other cells in this worksheet. Start by entering data in cells **A1**, **A2**, and **A3**, such as:

    **A1**: This text is exciting!
    **A2**: This text is boring.
    **A3**: This text is ordinary.

    Next, ask the model to evaluate the sentiments of the statements in cells **A1:A3**. Use the & character to append a cell address to the prompt. In cell **B1**, enter:

    ```
    =LABS.GENERATIVEAI("Analyze the sentiment of the following text:" &A1)
    ```

    After it returns the result in cell **B1**, copy the formula by dragging the fill handle down over cells **B2** and **B3**. You’ll get results from the model for each cell in the range **A1:A3**.

1. Answer a factual question. Enter "France" in cell **A1**. In another cell, enter:

    ```
    =LABS.GENERATIVEAI("What is the capital city of: " &A1,0)
    ```

1. Produce a response based on a sample. Teach the generative AI model how to format its output by providing an example in the prompt. Set the `temperature` value to 0.1 to limit creativity.

    Enter "Circle" in cell **A1**. In another cell, enter:

    ```
    =LABS.GENERATIVEAI("Q: Square. A: Cube. Q: " & A2 & ". A: ", 0.1)
    ```

1. Use the TEXTSPLIT formula with `LABS.GENERATIVEAI` to create a table of data that spills into multiple cells. Set `temperature` to 0.1 for limited creativity and set `max_tokens` to 1000 to extend the output length.

    ```
    =TEXTSPLIT(LABS.GENERATIVEAI("Return a table of helpful Excel " & "functions, with 3 columns: " & "name, category, and an example",0,1000), "|",UNICHAR(10),TRUE,,"")
    ```
    
    Using a prompt such as "as a table" is a reliable way to make the model return tabular data separated by the vertical bar character, “|”.
