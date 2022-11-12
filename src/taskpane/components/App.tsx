import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
/* global Word */


// https://learn.microsoft.com/en-us/javascript/api/
// npm install -g ts-node@latest

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  mode: number;
  message: string;
  scope: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      mode: 0,
      listItems: [],
      message: '',
      scope: ''
    };
  }

  botId = 'dev-rodriguez';
  botKey = 'starter';
  host = 'http://localhost';
  breakpointsMap = {};

  componentDidMount() {

    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Office integration to Bots",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features of General Bots",
        },
        {
          icon: "Design",
          primaryText: "Create your Bots using BASIC",
        },
      ],
    });

    setTimeout((() => {
      const context = this.updateExecutionContext();

      this.setState({
        mode: context['state'],
        message: context['messages'][0],
        scope: context['scope']
      });


    }).bind(this), 3000);

  }

  updateExecutionContext = async () => {

    const url = `${this.host}/debugger/${this.botId}/getExecutionContext`;

    $.ajax({
      data: { botId: this.botId, botKey: this.botKey, line },
      url: url,
      dataType: 'json',
    }).done(function (item) {

      Word.run(async (context) => {
        var paragraphs = context.document.body.paragraphs;
        paragraphs.load("$none");
        await context.sync();
        for (let i = 0; i < paragraphs.items.length; i++) {
          const paragraph = paragraphs.items[i]

          context.load(paragraph, ["text", "font"]);
          paragraph.font.highlightColor = null;

          if (i === line) {
            paragraph.font.highlightColor = 'Yellow';
          }
        }
        await context.sync();
      });

    }).fail(function (error) {
      console.log(error);
    });
  };


  setExecutionLine = async (line) => {
    Word.run(async (context) => {
      var paragraphs = context.document.body.paragraphs;
      paragraphs.load("$none");
      await context.sync();
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i]

        context.load(paragraph, ["text", "font"]);
        paragraph.font.highlightColor = null;

        if (i === line) {
          paragraph.font.highlightColor = 'Yellow';
        }
      }
      await context.sync();
    });
  };

  setBreakpoint = async () => {

    let line = 0;

    Word.run(async (context) => {

      let selection = context.document.getSelection();
      selection.load();

      await context.sync();

      console.log("Empty selection, cursor.");

      const paragraph = selection.paragraphs.getFirst();
      paragraph.select();
      context.load(paragraph, ["text", "font"]);

      var paragraphs = context.document.body.paragraphs;
      paragraphs.load("$none");
      await context.sync();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph1 = paragraphs.items[i]

        if (paragraph1 === paragraph) {
          line = i + 1;
          paragraph.font.highlightColor = 'Orange';
        }
      }

      return context.sync();
    });


    const url = `${this.host}/debugger/${this.botId}/setBreakpoint`;

    $.ajax({
      data: { botId: this.botId, botKey: this.botKey, line },
      url: url,
      dataType: 'json',
    }).done(function (item) {

    }).fail(function (error) {
      console.log(error);
    });
  }

  continueExecution = async () => {

    const url = `${this.host}/debugger/${this.botId}/continueRun`;

    $.ajax({
      data: { botId: this.botId, botKey: this.botKey },
      url: url,
      dataType: 'json',
    }).done(function (item) {
      this.state.mode = 1;
    }).fail(function (error) {
      console.log(error);
    });
  }



  stepOver = async () => {

    const url = `${this.host}/debugger/${this.botId}/stepOver`;

    $.ajax({
      data: { botId: this.botId, botKey: this.botKey },
      url: url,
      dataType: 'json',
    }).done(function (item) {
      this.state.mode = 2;
    }).fail(function (error) {
      console.log(error);
    });
  }


  stop = async () => {

    const url = `${this.host}/debugger/${this.botId}/stop`;

    $.ajax({
      data: { botId: this.botId, botKey: this.botKey },
      url: url,
      dataType: 'json',
    }).done(function (item) {
      this.state.mode = 0;
    }).fail(function (error) {
      console.log(error);
    });
  }


  run = async () => {

    if (this.state.mode === 0) {

      const url = `${this.host}/debugger/${this.botId}/run`;

      $.ajax({
        data: { botId: this.botId, botKey: this.botKey },
        url: url,
        dataType: 'json',
      }).done(function (item) {
        this.state.mode = 1;
      }).fail(function (error) {
        console.log(error);
      });
    }
    else {
      this.continueExecution();
    }
  }

  formatCode = async () => {
    return Word.run(async (context) => {

      var paragraphs = context.document.body.paragraphs;
      paragraphs.load("$none");
      await context.sync();
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i]
        context.load(paragraph, ["text", "font"]);
        paragraph.font.highlightColor = null;

        var words = paragraph.getTextRanges([" "], true);
        context.load(words, ["text", "font"]);
        var boldWords = [];
        for (var j = 0; j < words.items.length; ++j) {
          var word = words.items[j];
          if (word.text === "TALK" && j == 0) boldWords.push(word);
          if (word.text === "HEAR" && j == 0) boldWords.push(word);
          if (word.text === "SAVE" && j == 0) boldWords.push(word);
          if (word.text === "FIND" && j == 3) boldWords.push(word);
          if (word.text === "OPEN" && j == 0) boldWords.push(word);
        }
        for (var j = 0; j < boldWords.length; ++j) {
          boldWords[j].font.color = "blue";
          boldWords[j].font.bold = true;
        }
      }
      await context.sync();


    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <div>Excution Mode: {this.state.mode} </div>
        <div>Bot Messages: {this.state.message} </div>
        <div>Debug Scope: {this.state.scope} </div>
        <HeroList message="Discover what General Bots can do for you today!" items={this.state.listItems}>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "MdFormatPaint" }}
            onClick={this.formatCode}
          >
            Format .gbdialog
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "BsPlayCircle" }}
            onClick={this.run}
          >
            Run
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "BsStopCircle" }}
            onClick={this.stop}
          >
            Stop
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "BsFillDashCircleFill" }}
            onClick={this.setBreakpoint}
          >
            Set Breakpoint
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "VscDebugStepOver" }}
            onClick={this.stepOver}
          >
            Step Over
          </Button>
        </HeroList>
      </div>
    );
  }
}
