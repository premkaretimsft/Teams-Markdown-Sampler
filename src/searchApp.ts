import {
  TeamsActivityHandler,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  InvokeResponse,
  AdaptiveCardInvokeResponse,
  ActivityTypes,
  Activity,
  CardFactory,
  Attachment,
  MessageFactory,
  ActionTypes,
  ChannelAccount,
} from "botbuilder";
import productSearchCommand from "./messageExtensions/productSearchCommand";
import discountedSearchCommand from "./messageExtensions/discountSearchCommand";
import revenueSearchCommand from "./messageExtensions/revenueSearchCommand";
import actionHandler from "./adaptiveCards/cardHandler";
import {
  CreateActionErrorResponse,
  CreateInvokeResponse,
} from "./adaptiveCards/utils";
import { setTimeout as nodeTimeout } from "timers/promises";

const streamingPacketsText = [
  "Prem streaming- Second informative request that is meant to be a little longer than the first one",
  `<p>In a quiet forest, a small stream flowed peacefully through the trees. It was no ordinary stream, for it was said to hold magical powers that could grant wishes to those who drank from its waters.<br/><br/>
One day, a young girl named Lily stumbled upon the stream while wandering through the forest. She was lost and had been wandering for hours, but when she saw the stream, an overwhelming feeling of hope filled her heart.</p>`,
  `<p>Desperate for a way out of the forest, Lily approached the stream and knelt down to take a drink. As she sipped the cool, clear water, she closed her eyes and made a wish with all her heart.<br/><br/>
Suddenly, the world around her began to spin and swirl. When she opened her eyes again, she was no longer in the forest. Instead, she found herself in a grand castle, with marble floors and glittering chandeliers hanging overhead.<br/><br/>
Confused but curious, Lily began to explore the castle. She wandered through grand ballrooms, ornate dining halls, and even a secret garden filled with rare flowers and exotic birds.</p>`,
  `<p>As she explored, she began to realize that the castle belonged to a powerful sorceress. She had heard stories of the sorceress before, of the endless riches and magical powers she possessed.<br/><br/> Lily realized that her wish had brought her to this place, and she knew that she had to find a way to make the most of this opportunity.
For days, Lily explored the castle, learning all she could about the sorceress and her powers. She watched as the sorceress performed spells and incantations, and slowly but surely, she began to learn the ways of magic.<br/><br/>
In time, Lily became a powerful sorceress in her own right. She learned to control the elements, to summon creatures from the forest, and to cast spells that could bend reality itself. And all because of a wish she made at a magical stream in the heart of a forest.</p>`,
  `<p>In time, Lily became a powerful sorceress in her own right. She learned to control the elements, to summon creatures from the forest, and to cast spells that could bend reality itself.<br/><br/> And all because of a wish she made at a magical stream in the heart of a forest.</p>`,
];

export class SearchApp extends TeamsActivityHandler {
  notifyContinuationActivity: any;
  continuationParameters: any;

  constructor(
    notifyContinuationActivity: any,
    continuationParameters: any = {} /*conversations: any = {}*/
  ) {
    super();
    this.notifyContinuationActivity = notifyContinuationActivity;
    this.continuationParameters = continuationParameters;
  }

  public async onMessageActivity(context: TurnContext): Promise<void> {
    // this.addOrUpdateChannelPostParameters(context);
    if (!context.activity.text.includes("error")) {
      await context.sendActivity(this.getBotAIGenActivity(context));
    } else {
      await context.sendActivity({
        type: "message",
        text: `Bot error simulation: ${context.activity.text}`,
        entities: [
          {
            type: "BotMessageMetadata",
            botErrorInfo: {
              errorMessage:
                "Simulated error occurred while processing the message activity.",
              errorCode: 999,
            },
          },
        ],
      });
      // throw new Error("Simulated error occurred while processing the message activity.");
    }
  }

  private getStreamingActivity(
    streamType: string,
    streamId: string,
    sequence: number,
    addAttachments: boolean = false
  ): Partial<Activity> {
    let textContent = streamingPacketsText[sequence];

    if (streamType !== "informative") {
      textContent = streamingPacketsText.slice(1, sequence + 1).join(`<br/>`);
    }

    const activity = {
      type:
        streamType === "final" ? ActivityTypes.Message : ActivityTypes.Typing,
      text: textContent,
      entities: [
        {
          type: "streaminfo",
          streamType: streamType,
          streamId: streamId,
        },
      ],
    };

    if (streamType !== "final") {
      activity.entities[0]["streamSequence"] = sequence + 2;
      // activity["attachments"] = [this.getChartAdaptiveCard()];
    } else {
      if (addAttachments) {
        activity["attachments"] = [this.getChartAdaptiveCard()];
      }
    }

    return activity;
  }

  private async processStreamingRequest(
    context: TurnContext,
    addAttachments: boolean = false,
    packetDelay: number = 900
  ): Promise<void> {
    const result = await context.sendActivity({
      type: ActivityTypes.Typing,
      text: "Prem streaming- first informative request that is not too long",
      entities: [
        {
          type: "streaminfo",
          streamType: "informative", // informative or streaming; default= streaming.
          streamSequence: 1, // (required) incremental integer; must be present for start and continue streaming request, but must not be set for final streaming request.
        },
      ],
    });

    const streamId = result.id;
    console.log(`streamId: ${streamId}`);

    const streamingPacketsLength = streamingPacketsText.length;
    for (let i = 0; i < streamingPacketsLength; i++) {
      await nodeTimeout(packetDelay);
      const streamType =
        i === 0
          ? "informative"
          : i === streamingPacketsLength - 1
          ? "final"
          : "streaming";
      await context.sendActivity(
        this.getStreamingActivity(streamType, streamId, i, addAttachments)
      );
    }
  }

  public async onTeamsMessageEdit(context: TurnContext): Promise<void> {
    const conversationReference = TurnContext.getConversationReference(
      context.activity
    );

    const isChannel =
      conversationReference.conversation.conversationType === "channel";
    await context.sendActivity(
      this.getBotAIGenActivity(context, isChannel, true)
    );
  }

  private getBotAIGenActivity(
    context: TurnContext,
    isChannelPost: boolean = false,
    isEdited: boolean = false
  ): Partial<Activity> {
    const activityText = context.activity.text;
    console.log(`${activityText} - ${isChannelPost} - ${isEdited}`);

    return {
      type: "message",
      text: `You said: ${activityText}`,
      suggestedActions: {
        to: [context.activity.from.id],
        actions: [
          {
            type: ActionTypes.ImBack,
            title: this.generateMarkdownTestContent(activityText),
            value: "markdown",
          },
        ],
      },
      // attachments: [this.getChartXValuesCheckCard()],
      entities: [
        {
          type: "BotMessageMetadata",
          aiMetadata: {
            botAiSkill: "NewMarkdownMessageType",
          },
        },
        {
          type: "https://schema.org/Message",
          "@type": "Message",
          "@context": "https://schema.org",
          additionalType: ["AIGeneratedContent"], // AI Generated label
        },
      ],
      channelData: {
        feedbackLoopEnabled: true, // Enable feedback buttons
      },
    };
  }

  /**
   * Generates comprehensive markdown test content for Teams client rendering validation.
   * Includes various markdown elements with proper spacing for thorough testing.
   * @param activityText - The user's input text to incorporate into examples
   * @returns Complete markdown string with extensive formatting examples
   */
  private generateMarkdownTestContent(activityText: string): string {
    return `**You said this in bold: ${activityText} - This text should appear bold and emphasized!**

_In italic: "The quick brown fox jumps over the lazy dog" - This text should appear slanted and stylized with elegant cursive-like rendering_

***In bold italic: Welcome to the Microsoft 365 Copilot Testing Arena! - This combines both bold AND italic formatting for maximum emphasis!***

~~In strikethrough: Outdated information from version 1.0.0 has been deprecated and crossed out~~

In link: Click here to visit [Microsoft Teams Developer Platform Documentation](https://docs.microsoft.com/microsoftteams/platform) for comprehensive guides

In code span: \`const apiEndpoint = "https://graph.microsoft.com/v1.0/me"; // inline code formatting with realistic API call\`

# 🚀 Main Heading: Advanced Markdown Rendering Test Suite v2.5.1

## 🎯 Subheading: React-Markdown Component Performance Analysis

### 📈 Sub-subheading: Real-time Rendering Statistics Dashboard

> **Albert Einstein once said:** "Imagination is more important than knowledge. Knowledge is limited, whereas imagination embraces the entire world, stimulating progress, giving birth to evolution."
> 
> This multi-line blockquote demonstrates how philosophical quotes should render with proper indentation, styling, and text flow across multiple lines in your Teams client.

**🛠️ Advanced Feature Testing Checklist:**

- 🔥 Process complex nested markdown structures
- ⚡ Handle special characters: @#$%^&*(){}[]|\\;':",./<>?
- 🎯 Render emojis and Unicode symbols: 🌟✨🎉🚀💡🔥⭐
- 🌍 Support international text: Héllo Wörld! 你好世界! مرحبا بالعالم!
- ✨ Format mathematical expressions: E=mc² and π≈3.14159

**🔬 Scientific Workflow Process:**

1. 📊 Initialize test environment with sample data set (n=1000)
2. 🧪 Execute controlled experiments with variables A, B, and C
3. 📈 Collect performance metrics: latency, throughput, error rates
4. 🔍 Analyze results using statistical significance tests (p<0.05)
5. 📝 Generate comprehensive reports with visualizations
6. 🚀 Deploy optimized solution to production environment

📊 **Performance Metrics Dashboard:**

Component | Render Time (ms) | Memory Usage (MB) | CPU Load (%) | Status | Optimization Score
--- | --- | --- | --- | --- | ---
Header Navigation | 12.5 | 2.3 | 1.2 | � Optimal | 95/100
Content Renderer | 45.8 | 8.7 | 4.5 | 🟡 Good | 87/100
Markdown Parser | 23.1 | 5.2 | 2.8 | 🟢 Excellent | 98/100
Image Processor | 156.3 | 15.9 | 12.7 | 🟠 Moderate | 73/100
Table Generator | 34.7 | 6.1 | 3.2 | 🟢 Great | 91/100
Syntax Highlighter | 67.2 | 11.4 | 7.8 | 🟢 Good | 85/100

**�💻 Advanced Code Example - Teams Bot Implementation:**

**🎨 Visual Separator with Custom Styling:**

---

**🌈 Multi-line Formatting Demonstration:**

🎭 **Current Status:** Testing react-markdown rendering capabilities
⚡ **Processing Speed:** 2,847 operations per second
🎯 **Accuracy Rate:** 99.7% successful markdown transformations
🔮 **Next Phase:** Advanced interactive component integration
✨ **Final Goal:** Seamless Teams client markdown experience!

**🖼️ Dynamic Test Image:** ![Complex Markdown Test Visualization](https://via.placeholder.com/400x200/FF6B6B/FFFFFF?text=React+Markdown+Test+Suite+v2.5.1)

**🚀 Comprehensive Project Roadmap:**

- [x] ✅ Initialize markdown parsing engine with TypeScript support
- [x] 🎨 Implement custom styling for Teams-specific components
- [x] 📱 Test responsive design across different screen sizes
- [x] 🔧 Configure webpack optimization for production builds
- [ ] 🧪 Conduct A/B testing with focus groups (target: 500 users)
- [ ] 🌍 Add internationalization support for 12+ languages
- [ ] 🔒 Implement advanced security measures and data encryption
- [ ] 📈 Deploy analytics tracking for user engagement metrics
- [ ] 🚀 Launch beta version to Microsoft Teams App Store
- [ ] 🏆 Achieve 4.8+ star rating and 10,000+ active installations

**💻 Advanced Code Example - Teams Bot Implementation:**

\`\`\`typescript
// Advanced Teams Bot with Adaptive Cards and Graph API integration
import { TeamsActivityHandler, CardFactory, MessageFactory } from 'botbuilder';
import { Client } from '@microsoft/microsoft-graph-client';

interface TeamsMember {
  id: string;
  displayName: string;
  email: string;
  roles: string[];
  lastActive: Date;
}

class AdvancedTeamsBot extends TeamsActivityHandler {
  private graphClient: Client;
  
  constructor(graphClient: Client) {
    super();
    this.graphClient = graphClient;
  }
  
  protected async onMessageActivity(context: TurnContext): Promise<void> {
    const userMessage = context.activity.text?.toLowerCase();
    
    if (userMessage?.includes('analytics')) {
      const analyticsCard = this.createAnalyticsCard();
      await context.sendActivity(MessageFactory.attachment(analyticsCard));
    }
  }
  
  private createAnalyticsCard(): Attachment {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      version: '1.4',
      body: [{
        type: 'TextBlock',
        text: 'Advanced Analytics Dashboard',
        weight: 'Bolder',
        size: 'Large'
      }]
    });
  }
}

export { AdvancedTeamsBot, TeamsMember };
\`\`\``;
  }

  private getChartXValuesCheckCard(): Attachment {
    return CardFactory.adaptiveCard({
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      msTeams: {
        width: "full",
      },
      body: [
        {
          type: "Chart.Line",
          colorSet: "categorical",
          data: [
            {
              color: "categoricalBlue",
              values: [
                {
                  x: "8Y",
                  y: 3.921,
                },
                {
                  x: "9Y",
                  y: 3.994,
                },
                {
                  x: "10Y",
                  y: 4.065,
                },
                {
                  x: "11Y",
                  y: 4.131,
                },
                {
                  x: "12Y",
                  y: 4.195,
                },
              ],
            },
          ],
        },
      ],
    });
  }

  private getChartInputActionsCard(): Attachment {
    return CardFactory.adaptiveCard({
      type: "AdaptiveCard",
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.6",
      body: [
        {
          type: "Input.Text",
          placeholder: "Placeholder text",
          label: "Text input",
          id: "text",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Date",
          label: "Date input",
          id: "date",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Time",
          id: "time",
          label: "Time input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Number",
          placeholder: "Placeholder text",
          id: "number",
          label: "Number input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.ChoiceSet",
          choices: [
            {
              title: "Choice 1",
              value: "Choice 1",
            },
            {
              title: "Choice 2",
              value: "Choice 2",
            },
          ],
          placeholder: "Placeholder text",
          id: "choiceSet",
          label: "ChoiceSet input",
          isRequired: true,
          errorMessage: "Error",
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Action.Submit",
          conditionallyEnabled: true,
          associatedInputs: "auto",
        },
        {
          type: "Action.OpenUrl",
          title: "Handoff to Bot",
          url: "https://teams.microsoft.com/l/chat/0/0?users=28:ff6df678-2ef6-47b9-b604-206d8686ea31&continuation=premtestconttoken",
        },
      ],
    });
  }

  private getChartAdaptiveCard(): Attachment {
    const groupedBarChartData = [
      {
        legend: "Outlook",
        values: [
          { x: "2023-05-01", y: 24 },
          { x: "2023-05-02", y: 27 },
          { x: "2023-05-03", y: 18 },
          { x: "2023-05-04", y: 30 },
          { x: "2023-05-05", y: 20 },
          { x: "2023-05-06", y: 35 },
          { x: "2023-05-07", y: 40 },
          { x: "2023-05-08", y: 45 },
        ],
      },
      {
        legend: "Teams",
        values: [
          { x: "2023-05-01", y: 9 },
          { x: "2023-05-02", y: 100 },
          { x: "2023-05-03", y: 22 },
          { x: "2023-05-04", y: 40 },
          { x: "2023-05-05", y: 30 },
          { x: "2023-05-06", y: 45 },
          { x: "2023-05-07", y: 50 },
          { x: "2023-05-08", y: 55 },
        ],
      },
      {
        legend: "Office",
        values: [
          { x: "2023-05-01", y: 10 },
          { x: "2023-05-02", y: 20 },
          { x: "2023-05-03", y: 30 },
          { x: "2023-05-04", y: 40 },
          { x: "2023-05-05", y: 50 },
          { x: "2023-05-06", y: 60 },
          { x: "2023-05-07", y: 70 },
          { x: "2023-05-08", y: 80 },
        ],
      },
      {
        legend: "Windows",
        values: [
          { x: "2023-05-01", y: 10 },
          { x: "2023-05-02", y: 20 },
          { x: "2023-05-03", y: 30 },
          { x: "2023-05-04", y: 40 },
          { x: "2023-05-05", y: 50 },
          { x: "2023-05-06", y: 60 },
          { x: "2023-05-07", y: 70 },
          { x: "2023-05-08", y: 80 },
        ],
      },
    ];

    return CardFactory.adaptiveCard({
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        {
          type: "TextBlock",
          text: "Simple",
          size: "large",
          separator: true,
          spacing: "large",
        },
        {
          type: "Chart.VerticalBar",
          title: "Sample",
          xAxisTitle: "Days",
          yAxisTitle: "Sales",
          colorSet: "categorical",
          data: [
            { x: "Pear", y: 59 },
            { x: "Banana", y: 292 },
            { x: "Apple", y: 143 },
            { x: "Peach", y: 98 },
            { x: "Kiwi", y: 179 },
            { x: "Grapefruit", y: 20 },
            { x: "Orange", y: 212 },
            { x: "Cantaloupe", y: 68 },
            { x: "Grape", y: 102 },
            { x: "Tangerine", y: 38 },
          ],
        },
        {
          type: "TextBlock",
          text: "Grouped",
          size: "large",
          separator: true,
          spacing: "large",
        },
        {
          type: "Chart.VerticalBar.Grouped",
          title: "Sample",
          xAxisTitle: "Days",
          yAxisTitle: "Sales",
          colorSet: "diverging",
          data: groupedBarChartData,
        },
        {
          type: "TextBlock",
          text: "Stacked",
          size: "large",
          separator: true,
          spacing: "large",
        },
        {
          type: "Chart.VerticalBar.Grouped",
          stacked: true,
          title: "Sample",
          xAxisTitle: "Days",
          yAxisTitle: "Sales",
          data: groupedBarChartData,
        },
        {
          type: "Input.Text",
          placeholder: "Placeholder text",
          label: "Text input",
          id: "text",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Date",
          label: "Date input",
          id: "date",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Time",
          id: "time",
          label: "Time input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.Number",
          placeholder: "Placeholder text",
          id: "number",
          label: "Number input",
          isRequired: true,
          errorMessage: "Error",
        },
        {
          type: "Input.ChoiceSet",
          choices: [
            {
              title: "Choice 1",
              value: "Choice 1",
            },
            {
              title: "Choice 2",
              value: "Choice 2",
            },
          ],
          placeholder: "Placeholder text",
          id: "choiceSet",
          label: "ChoiceSet input",
          isRequired: true,
          errorMessage: "Error",
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Action.Submit",
          conditionallyEnabled: true,
          associatedInputs: "auto",
        },
      ],
    });
  }

  private addOrUpdateChannelPostParameters(context): void {
    const conversationReference = TurnContext.getConversationReference(
      context.activity
    );

    if (conversationReference.conversation.conversationType === "channel") {
      console.log(
        `Adding continuation parameters for channel context: ${JSON.stringify(
          context
        )}`
      );
      this.continuationParameters[context.activity.from.id] = {
        claimsIdentity: context.turnState.get(context.adapter.BotIdentityKey),
        conversationReference: TurnContext.getConversationReference(
          context.activity
        ),
        oAuthScope: context.turnState.get(context.adapter.OAuthScopeKey),
        partialActivity: this.getBotAIGenActivity(context, true),
      };

      setTimeout(async () => await this.notifyContinuationActivity(), 1000);
    }
  }

  public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    try {
      switch (context.activity.name) {
        case "odsl/statementExecuted":
          return CreateInvokeResponse(200);
        case "handoff/action":
          // return CreateInvokeResponse(200);
          return {
            status: 200,
            body: `handoff invoke received- ${context.activity.name}`,
          };
        case "message/submitAction":
          return CreateInvokeResponse(200);
        case "composeExtension/query":
          return {
            status: 200,
            body: await this.handleTeamsMessagingExtensionQuery(
              context,
              context.activity.value
            ),
          };
        case "adaptiveCard/action":
          return {
            status: 200,
            body: await this.onAdaptiveCardInvoke(context),
          };
        default:
          return {
            status: 200,
            body: `Unknown invoke activity handled as default- ${context.activity.name}`,
          };
      }
    } catch (err) {
      console.log(`Error in onInvokeActivity: ${err}`);
      return {
        status: 500,
        body: `Invoke activity received- ${context.activity.name}`,
      };
    }
  }

  // Handle search message extension
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    switch (query.commandId) {
      case productSearchCommand.COMMAND_ID: {
        return productSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      }
      case discountedSearchCommand.COMMAND_ID: {
        return discountedSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      }
      case revenueSearchCommand.COMMAND_ID: {
        return revenueSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      }
    }
  }

  // Handle adaptive card actions
  public async onAdaptiveCardInvoke(
    context: TurnContext
  ): Promise<AdaptiveCardInvokeResponse> {
    try {
      if (context.activity?.value?.action?.verb) {
        switch (context.activity.value.action.verb) {
          case "ok": {
            return actionHandler.handleTeamsCardActionUpdateStock(context);
          }
          case "restock": {
            return actionHandler.handleTeamsCardActionRestock(context);
          }
          case "cancel": {
            return actionHandler.handleTeamsCardActionCancelRestock(context);
          }
          default:
            return CreateActionErrorResponse(
              400,
              0,
              `ActionVerbNotSupported: ${context.activity.value.action.verb} is not a supported action verb.`
            );
        }
      } else {
        return {
          statusCode: 200,
          type: "application/vnd.microsoft.card.adaptive",
          value: {
            type: "AdaptiveCard",
            $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
            version: "1.6",
            body: [
              {
                type: "Input.ChoiceSet",
                choices: [
                  {
                    title: "Choice 1",
                    value: "Choice 1",
                  },
                  {
                    title: "Choice 2",
                    value: "Choice 2",
                  },
                ],
                placeholder: "Placeholder text",
                id: "choiceSet",
                label: "ChoiceSet input",
                isRequired: true,
                errorMessage: "Error",
              },
            ],
            actions: [
              {
                type: "Action.OpenUrl",
                title: "Invoke successful",
                url: "https://www.microsoft.com",
              },
            ],
          },
        };
      }
    } catch (err) {
      return CreateActionErrorResponse(500, 0, err.message);
    }
  }
}
