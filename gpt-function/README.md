# GPT function, an Excel Labs feature

GPT function creates a custom function called LABS.GPT that you can use to send prompts from the Excel grid to OpenAI’s GPT language model. LABS.GPT returns the results from GPT back to your worksheet.

You can send GPT simple or complex prompts to analyze information, import public data, produce a response based on a sample, and more. LABS.GPT allows you to reference other cells in your workbook, and it can be called inside any Excel cell or named formula in the workbook.

The GPT function also includes model settings, which let you to choose from six different GPT models, control the default output length of results returned from GPT, and more.

## Examples

The following examples show several ways to use LABS.GPT.

- Analyze public information. Ask GPT to parse the tweets from a particular account, such as:

    ```
    =LABS.GPT("What is the most common word in tweets from Microsoft?")
    ```

- Import data and define the format. Ask GPT to import a particular type of publicly available data, and tell GPT how to format the information, such as:

    ```
    =LABS.GPT("Return the airport codes for 2 airports. Return this as a comma-separated list in the format of: Full airport name, Country, Airport code")
    ```

- Answer a factual question. Reference other cells in this workbook to formulate the prompt:

    ```
    // Enter "France" in cell A1. In another cell, enter:
    =LABS.GPT("What is the capital city of: " &A1)
    ```

For more examples, see [GPT function samples](EXAMPLES.md).

## Requirements

The GPT function uses the GPT-3 (Generative Pre-trained Transformer 3) technology from OpenAI. To use the LABS.GPT function, you must register an OpenAI account and generate a unique API key.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.
