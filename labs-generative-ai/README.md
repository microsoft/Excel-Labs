# LABS.GENERATIVEAI function, an Excel Labs feature

`LABS.GENERATIVEAI` is custom function that allows you to send prompts from the Excel grid to a generative AI model and then return the results from the model directly back to your worksheet.

You can send the model simple or complex prompts to analyze information, process data, produce a response based on a sample, and more. `LABS.GENERATIVEAI` allows you to reference other cells in your workbook, and it can be called inside any Excel cell or named formula in the workbook.

The custom function includes settings which let you choose from different models, control the default output length of results returned from the models, limit creativity, and more.

## Getting started

To experiment with generative AI, start by typing the formula `=LABS.GENERATIVEAI` in a cell. Next, give the formula a prompt to send to OpenAI. Make sure to wrap your prompt in quotation marks, like `=LABS.GENERATIVEAI(“Write a poem about Excel.”)`, and select the Enter key to send the request to OpenAI.

Use the **Settings** options in the task pane to customize your experience. You can choose from different models, control the default output length of results returned from the models, and more.

## Formula syntax

In addition to the prompt argument, `LABS.GENERATIVEAI` offers the optional arguments `temperature`, `max_tokens`, and `model`. The formula uses the following syntax.

```
LABS.GENERATIVEAI(prompt, [temperature], [max_tokens], [model])
```

These optional arguments correspond to options in the **Settings** section. Use the **Settings** to control default settings across the workbook. Use the optional arguments in an individual `LABS.GENERATIVEAI` call override your default settings.

## Examples

The following examples show a couple ways to use `LABS.GENERATIVEAI`.

- Analyze public information. Ask a model to summarize a complicated topic, such as:

    ```
    =LABS.GENERATIVEAI("Explain generative AI in a sentence.”)
    ```

- Return creative content. Ask a model to create content about a particular subject.

    ```
    =LABS.GENERATIVEAI(“Write a poem about Excel.”)
    ```

For more examples, see [LABS.GENERATIVEAI function samples](EXAMPLES.md).

## Requirements

The `LABS.GENERATIVEAI` function uses generative AI technology from OpenAI. To use the function, you must register an OpenAI account and generate a unique API key.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.
