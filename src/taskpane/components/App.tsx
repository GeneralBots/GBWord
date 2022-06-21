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

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

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
  }

  click = async () => {
    return Word.run(async (context) => {
      var words = context.document.getSelection().getTextRanges([" "], true);
      context.load(words, ["text", "font"]);
      var boldRanges = [];
      return context
        .sync()
        .then(function () {
          for (var i = 0; i < words.items.length; ++i) {
            var word = words.items[i];
            if (word.text === "TALK") boldRanges.push(word);
          }
        })
        .then(function () {
          for (var j = 0; j < boldRanges.length; ++j) {
            boldRanges[j].font.color = "blue";
          }
        });
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
        <HeroList message="Discover what General Bots can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            <b>Format your Bot Code with a click</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Format .gbdialog
          </Button>
        </HeroList>
      </div>
    );
  }
}
