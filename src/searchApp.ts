import {
  TeamsActivityHandler,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  InvokeResponse,
  AdaptiveCardInvokeResponse,
  Activity,
  CardFactory,
  Attachment,
  ActionTypes,
  ActivityTypes,
} from "botbuilder";
import { setTimeout as nodeTimeout } from "timers/promises";
import productSearchCommand from "./messageExtensions/productSearchCommand";
import discountedSearchCommand from "./messageExtensions/discountSearchCommand";
import revenueSearchCommand from "./messageExtensions/revenueSearchCommand";
import actionHandler from "./adaptiveCards/cardHandler";
import {
  CreateActionErrorResponse,
  CreateInvokeResponse,
} from "./adaptiveCards/utils";

export class SearchApp extends TeamsActivityHandler {
  notifyContinuationActivity: any;
  continuationParameters: any;

  private markdownScenarios = new Map<string, string>([
    [
      "Basic Formatting",
      `**Bold Text Example:** This text should appear bold and emphasized!

_Italic Text Example:_ This text should appear slanted and stylized with elegant cursive-like rendering

***Bold Italic Combined:*** This combines both bold AND italic formatting for maximum emphasis!

~~Strikethrough Example:~~ This text should appear crossed out and deprecated`,
    ],

    [
      "Links and Code",
      `**Link Example:** Click here to visit [Microsoft Teams Developer Platform Documentation](https://docs.microsoft.com/microsoftteams/platform) for comprehensive guides

**Code Span Example:** \`const apiEndpoint = "https://graph.microsoft.com/v1.0/me"; // inline code formatting\`

**Special Characters:** Handle special characters: @#$%^&*(){}[]|\\;':",./<>?`,
    ],

    [
      "Headings and Structure",
      `# 🚀 Main Heading: Advanced Markdown Rendering Test Suite

## 🎯 Subheading: React-Markdown Component Performance Analysis  

### 📈 Sub-subheading: Real-time Rendering Statistics Dashboard

> **Quote Example:** "Imagination is more important than knowledge. Knowledge is limited, whereas imagination embraces the entire world, stimulating progress, giving birth to evolution." - Albert Einstein
> 
> This multi-line blockquote demonstrates proper indentation and styling.`,
    ],

    [
      "Lists and Emojis",
      `**🛠️ Feature Testing Checklist:**

- 🔥 Process complex nested markdown structures
- ⚡ Handle special characters and symbols
- 🎯 Render emojis and Unicode: 🌟✨🎉🚀💡🔥⭐
- 🌍 Support international text: Héllo Wörld! 你好世界! مرحبا بالعالم!
- ✨ Format mathematical expressions: E=mc² and π≈3.14159

**🔬 Numbered Process:**

1. 📊 Initialize test environment with sample data
2. 🧪 Execute controlled experiments
3. 📈 Collect performance metrics
4. 🔍 Analyze results using statistical tests
5. 📝 Generate comprehensive reports`,
    ],

    [
      "Performance Table",
      `📊 **Performance Metrics Dashboard:**

Component | Render Time (ms) | Memory Usage (MB) | CPU Load (%) | Status | Optimization Score
--- | --- | --- | --- | --- | ---
Header Navigation | 12.5 | 2.3 | 1.2 | 🟢 Optimal | 95/100
Content Renderer | 45.8 | 8.7 | 4.5 | 🟡 Good | 87/100
Markdown Parser | 23.1 | 5.2 | 2.8 | 🟢 Excellent | 98/100
Image Processor | 156.3 | 15.9 | 12.7 | 🟠 Moderate | 73/100
Table Generator | 34.7 | 6.1 | 3.2 | 🟢 Great | 91/100
Syntax Highlighter | 67.2 | 11.4 | 7.8 | 🟢 Good | 85/100`,
    ],

    [
      "Adaptive Card",
      `**🎴 Interactive Adaptive Card Examples:**

This example demonstrates how Teams can render rich, interactive Adaptive Cards using the \`adaptivecard\` fenced block syntax:

**Card Example 1: Project Schema Card**

\`\`\`adaptivecard
{
  "type": "AdaptiveCard",
  "version": "1.6",
  "body": [
    {
      "type": "TextBlock",
      "size": "Medium",
      "weight": "Bolder",
      "text": "Publish Adaptive Card schema"
    },
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "items": [
            {
              "type": "Image",
              "style": "Person",
              "url": "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
              "size": "Small"
            }
          ],
          "width": "auto"
        },
        {
          "type": "Column",
          "items": [
            {
              "type": "TextBlock",
              "weight": "Bolder",
              "text": "Matt Hidinger",
              "wrap": true
            },
            {
              "type": "TextBlock",
              "spacing": "None",
              "text": "Created {{DATE(2017-02-14T06:08:39Z,SHORT)}}",
              "isSubtle": true,
              "wrap": true
            }
          ],
          "width": "stretch"
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        {
          "title": "Board:",
          "value": "Adaptive Card"
        },
        {
          "title": "List:",
          "value": "Backlog"
        },
        {
          "title": "Assigned to:",
          "value": "Matt Hidinger"
        },
        {
          "title": "Due date:",
          "value": "Not set"
        }
      ]
    }
  ],
  "actions": [
    {
      "type": "Action.ShowCard",
      "title": "Set due date",
      "card": {
        "type": "AdaptiveCard",
        "body": [
          {
            "type": "Input.Date",
            "id": "dueDate"
          },
          {
            "type": "Input.Text",
            "id": "comment",
            "placeholder": "Add a comment",
            "isMultiline": true
          }
        ],
        "actions": [
          {
            "type": "Action.Submit",
            "title": "OK"
          }
        ]
      }
    },
    {
      "type": "Action.ShowCard",
      "title": "Comment",
      "card": {
        "type": "AdaptiveCard",
        "body": [
          {
            "type": "Input.Text",
            "id": "comment",
            "isMultiline": true,
            "placeholder": "Add a comment"
          }
        ],
        "actions": [
          {
            "type": "Action.Submit",
            "title": "OK"
          }
        ]
      }
    }
  ]
}
\`\`\`

**Card Example 2: Project Schema Card (Copy)**

\`\`\`adaptivecard
{
  "type": "AdaptiveCard",
  "version": "1.6",
  "body": [
    {
      "type": "TextBlock",
      "size": "Medium",
      "weight": "Bolder",
      "text": "Publish Adaptive Card schema"
    },
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "items": [
            {
              "type": "Image",
              "style": "Person",
              "url": "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
              "size": "Small"
            }
          ],
          "width": "auto"
        },
        {
          "type": "Column",
          "items": [
            {
              "type": "TextBlock",
              "weight": "Bolder",
              "text": "Matt Hidinger",
              "wrap": true
            },
            {
              "type": "TextBlock",
              "spacing": "None",
              "text": "Created {{DATE(2017-02-14T06:08:39Z,SHORT)}}",
              "isSubtle": true,
              "wrap": true
            }
          ],
          "width": "stretch"
        }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        {
          "title": "Board:",
          "value": "Adaptive Card"
        },
        {
          "title": "List:",
          "value": "Backlog"
        },
        {
          "title": "Assigned to:",
          "value": "Matt Hidinger"
        },
        {
          "title": "Due date:",
          "value": "Not set"
        }
      ]
    }
  ],
  "actions": [
    {
      "type": "Action.ShowCard",
      "title": "Set due date",
      "card": {
        "type": "AdaptiveCard",
        "body": [
          {
            "type": "Input.Date",
            "id": "dueDate"
          },
          {
            "type": "Input.Text",
            "id": "comment",
            "placeholder": "Add a comment",
            "isMultiline": true
          }
        ],
        "actions": [
          {
            "type": "Action.Submit",
            "title": "OK"
          }
        ]
      }
    },
    {
      "type": "Action.ShowCard",
      "title": "Comment",
      "card": {
        "type": "AdaptiveCard",
        "body": [
          {
            "type": "Input.Text",
            "id": "comment",
            "isMultiline": true,
            "placeholder": "Add a comment"
          }
        ],
        "actions": [
          {
            "type": "Action.Submit",
            "title": "OK"
          }
        ]
      }
    }
  ]
}
\`\`\`

**Features demonstrated:**
- TextBlock with styling options and colors
- ColumnSet layout with Person images
- FactSet for structured data display
- Action.ShowCard for interactive elements
- Nested Adaptive Cards with form inputs
- ChoiceSet with multiple options
- Multiple card variations within one scenario`,
    ],

    [
      "Project Roadmap",
      `**🚀 Comprehensive Project Roadmap:**

- [x] ✅ Initialize markdown parsing engine with TypeScript support
- [x] 🎨 Implement custom styling for Teams-specific components  
- [x] 📱 Test responsive design across different screen sizes
- [x] 🔧 Configure webpack optimization for production builds
- [ ] 🧪 Conduct A/B testing with focus groups (target: 500 users)
- [ ] 🌍 Add internationalization support for 12+ languages
- [ ] 🔒 Implement advanced security measures and data encryption
- [ ] 📈 Deploy analytics tracking for user engagement metrics
- [ ] 🚀 Launch beta version to Microsoft Teams App Store
- [ ] 🏆 Achieve 4.8+ star rating and 10,000+ active installations`,
    ],

    [
      "TypeScript Code",
      `**💻 Advanced Code Example - Teams Bot Implementation:**

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
\`\`\``,
    ],

    [
      "KaTeX Math Equations",
      `**🧮 Mathematical Expressions with KaTeX:**

**Quadratic Formula:** The solution to $ax^2 + bx + c = 0$ is given by:

$$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$

**Euler's Identity:** One of the most beautiful equations in mathematics:

$$e^{i\\pi} + 1 = 0$$

**Fourier Transform:** Converting from time domain to frequency domain:

$$F(\\omega) = \\int_{-\\infty}^{\\infty} f(t) e^{-i\\omega t} dt$$

**Inline Math Examples:** 
- Area of circle: $A = \\pi r^2$ 
- Pythagorean theorem: $a^2 + b^2 = c^2$
- Derivative: $\\frac{d}{dx}(x^n) = nx^{n-1}$

**Matrix Operations:**

$$\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix} \\begin{pmatrix} x \\\\ y \\end{pmatrix} = \\begin{pmatrix} ax + by \\\\ cx + dy \\end{pmatrix}$$`,
    ],

    [
      "Visual Elements",
      `**� Visual Separator with Custom Styling:**

---

**🌈 Multi-line Formatting Demonstration:**

🎭 **Current Status:** Testing react-markdown rendering capabilities
⚡ **Processing Speed:** 2,847 operations per second  
🎯 **Accuracy Rate:** 99.7% successful markdown transformations
🔮 **Next Phase:** Advanced interactive component integration
✨ **Final Goal:** Seamless Teams client markdown experience!

**🖼️ Dynamic Test Image:** ![Markdown Test Visualization](https://via.placeholder.com/400x200/FF6B6B/FFFFFF?text=React+Markdown+Test+Suite)`,
    ],
  ]);

  constructor(
    notifyContinuationActivity: any,
    continuationParameters: any = {} /*conversations: any = {}*/
  ) {
    super();
    this.notifyContinuationActivity = notifyContinuationActivity;
    this.continuationParameters = continuationParameters;

    // Add the combined "All" scenario after construction
    this.markdownScenarios.set("All", this.getCombinedMarkdownScenarios());
  }

  private getCombinedMarkdownScenarios(): string {
    const allScenarios = Array.from(this.markdownScenarios.entries())
      .filter(([key]) => key !== "All") // Exclude "All" to avoid circular reference
      .map(([title, content]) => `# 📋 ${title}\n\n${content}`)
      .join("\n\n---\n\n");

    return `# 🎯 Complete Markdown Test Suite - All Scenarios Combined\n\n${allScenarios}`;
  }

  private getStreamingActivity(
    streamType: string,
    streamId: string,
    sequence: number,
    selectedScenario: string,
    markdownContent: string,
    fromId: string
  ): Partial<Activity> {
    let textContent: string;

    if (streamType === "informative") {
      textContent = sequence === 0 
        ? `🔍 Preparing markdown scenario: ${selectedScenario}...`
        : `🎨 Rendering markdown content with enhanced formatting...`;
    } else {
      textContent = `You selected: ${selectedScenario}`;
    }

    const activity: Partial<Activity> = {
      type: streamType === "final" ? ActivityTypes.Message : ActivityTypes.Typing,
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
    } else {
      activity.suggestedActions = {
        to: [fromId],
        actions: [
          {
            type: ActionTypes.ImBack,
            title: "Tag",
            value: "NewMarkdownMessageType",
          },
          {
            type: ActionTypes.ImBack,
            title: markdownContent,
            value: "markdown",
          },
        ],
      };
    }

    return activity;
  }

  private async processMarkdownStreamingRequest(
    context: TurnContext,
    selectedScenario: string,
    markdownContent: string
  ): Promise<void> {
    const result = await context.sendActivity({
      type: ActivityTypes.Typing,
      text: `🔍 Preparing markdown scenario: ${selectedScenario}...`,
      entities: [
        {
          type: "streaminfo",
          streamType: "informative",
          streamSequence: 1,
        },
      ],
    });

    const streamId = result.id;
    console.log(`streamId: ${streamId}`);

    await nodeTimeout(1000);

    const secondActivity = this.getStreamingActivity(
      "informative",
      streamId,
      1,
      selectedScenario,
      markdownContent,
      context.activity.from.id
    );
    
    await context.sendActivity(secondActivity);

    await nodeTimeout(1500);

    const finalActivity = this.getStreamingActivity(
      "final",
      streamId,
      2,
      selectedScenario,
      markdownContent,
      context.activity.from.id
    );
    
    await context.sendActivity(finalActivity);
  }

  public async onMessageActivity(context: TurnContext): Promise<void> {
    // this.addOrUpdateChannelPostParameters(context);
    const userText = context.activity.text?.toLowerCase().trim();

    // Handle 'list' command to show available scenarios
    if (userText === "list") {
      const scenarioList = Array.from(this.markdownScenarios.keys())
        .map((scenario, index) => `${index + 1}. ${scenario}`)
        .join("\n");

      await context.sendActivity({
        type: "message",
        text: `Available Markdown Test Scenarios:\n\n${scenarioList}\n\nReply with a number (1-${this.markdownScenarios.size}) to see that scenario.`,
      });
      return;
    }

    // Handle numeric responses for scenario selection
    const scenarioNumber = parseInt(userText || "");
    if (
      !isNaN(scenarioNumber) &&
      scenarioNumber >= 1 &&
      scenarioNumber <= this.markdownScenarios.size
    ) {
      const scenarioNames = Array.from(this.markdownScenarios.keys());
      const selectedScenario = scenarioNames[scenarioNumber - 1];
      const markdownContent = this.markdownScenarios.get(selectedScenario);

      // Use streaming approach for markdown scenario selection
      await this.processMarkdownStreamingRequest(context, selectedScenario, markdownContent);
      return;
    }

    // Default behavior for other messages
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

  private getBotAIGenActivity(
    context: TurnContext,
    isChannelPost: boolean = false,
    isEdited: boolean = false
  ): Partial<Activity> {
    const activityText = context.activity.text;
    console.log(`${activityText} - ${isChannelPost} - ${isEdited}`);

    return {
      type: "message",
      text: `You said: ${activityText}\n\n💡 **Tip:** Type 'list' to see available markdown test scenarios, or type a number (1-${this.markdownScenarios.size}) to select a specific scenario.`,
      entities: [
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

  private getChartAdaptiveCard(): Attachment {
    return CardFactory.adaptiveCard({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.6",
      body: [
        {
          type: "TextBlock",
          size: "Medium",
          weight: "Bolder",
          text: "Publish Adaptive Card schema",
        },
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              items: [
                {
                  type: "Image",
                  style: "Person",
                  url: "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
                  size: "Small",
                },
              ],
              width: "auto",
            },
            {
              type: "Column",
              items: [
                {
                  type: "TextBlock",
                  weight: "Bolder",
                  text: "Matt Hidinger",
                  wrap: true,
                },
                {
                  type: "TextBlock",
                  spacing: "None",
                  text: "Created {{DATE(2017-02-14T06:08:39Z,SHORT)}}",
                  isSubtle: true,
                  wrap: true,
                },
              ],
              width: "stretch",
            },
          ],
        },
        {
          type: "TextBlock",
          text: "Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.",
          wrap: true,
        },
        {
          type: "FactSet",
          facts: [
            {
              title: "Board:",
              value: "Adaptive Card",
            },
            {
              title: "List:",
              value: "Backlog",
            },
            {
              title: "Assigned to:",
              value: "Matt Hidinger",
            },
            {
              title: "Due date:",
              value: "Not set",
            },
          ],
        },
      ],
      actions: [
        {
          type: "Action.ShowCard",
          title: "Set due date",
          card: {
            type: "AdaptiveCard",
            body: [
              {
                type: "Input.Date",
                id: "dueDate",
              },
              {
                type: "Input.Text",
                id: "comment",
                placeholder: "Add a comment",
                isMultiline: true,
              },
            ],
            actions: [
              {
                type: "Action.Submit",
                title: "OK",
              },
            ],
          },
        },
        {
          type: "Action.ShowCard",
          title: "Comment",
          card: {
            type: "AdaptiveCard",
            body: [
              {
                type: "Input.Text",
                id: "comment",
                isMultiline: true,
                placeholder: "Add a comment",
              },
            ],
            actions: [
              {
                type: "Action.Submit",
                title: "OK",
              },
            ],
          },
        },
      ],
    });
  }

  public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    try {
      switch (context.activity.name) {
        case "odsl/statementExecuted":
          return CreateInvokeResponse(200, { type: "continue" });
        case "handoff/action":
          return CreateInvokeResponse(200, {
            handoff: { state: "completed" },
          });
        case "message/submitAction":
          return CreateInvokeResponse(200, {});
        case "composeExtension/query":
          const messagingExtensionResponse =
            await this.handleTeamsMessagingExtensionQuery(
              context,
              context.activity.value
            );
          return CreateInvokeResponse(200, messagingExtensionResponse);
        case "adaptiveCard/action":
          const adaptiveCardInvokeResponse: AdaptiveCardInvokeResponse =
            await actionHandler.handleTeamsCardActionUpdateStock(context);
          return CreateInvokeResponse(
            adaptiveCardInvokeResponse.statusCode,
            adaptiveCardInvokeResponse.value
          );
        default:
          const errorResponse = CreateActionErrorResponse(
            501,
            0,
            `NotImplemented: No handler for name ${context.activity.name}`
          );
          return CreateInvokeResponse(
            errorResponse.statusCode,
            errorResponse.value
          );
      }
    } catch (error) {
      console.log(`Error in onInvokeActivity: ${error}`);
      const errorResponse = CreateActionErrorResponse(500, 0, error.message);
      return CreateInvokeResponse(
        errorResponse.statusCode,
        errorResponse.value
      );
    }
  }

  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    switch (query.commandId) {
      case productSearchCommand.COMMAND_ID:
        return productSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      case discountedSearchCommand.COMMAND_ID:
        return discountedSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
      case revenueSearchCommand.COMMAND_ID:
        return revenueSearchCommand.handleTeamsMessagingExtensionQuery(
          context,
          query
        );
    }
  }
}
