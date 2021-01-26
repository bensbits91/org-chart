import * as React from 'react';
// import styles from './SharePointOrgChart.module.scss';
import { ISharePointOrgChartProps } from './ISharePointOrgChartProps';
// import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

// import OrgChart from './OrgChart';
import TheChart from './Chart';
// import InfoPanel from './Panel';

// import TempChart from './TempChart';
// import OrChar from './OrChar';

// import './temp.css';

// import ErrorBoundary from './ErrorBoundary';

const mcc = 'color:lime;background-color:black;';
const mcc2 = 'color:green;background-color:black;';

interface ISharePointOrgChartState {
  hierarchy_data?: any;
  all_data?: any;
  // panel_id?: string;
  panel_data?: any;
}

export default class SharePointOrgChart extends React.Component<ISharePointOrgChartProps, ISharePointOrgChartState> {

  public org_depth = 0;

  public constructor(props) {
    super(props);
    this.state = {
      // hierarchy_data: null
    };
    this.handler_card = this.handler_card.bind(this);
    this.handler_panel = this.handler_panel.bind(this);
  }

  public componentDidMount() {
    console.clear();

    this.get_siteUsers().then((su: []) => {
      console.log('%c : SharePointOrgChart -> componentDidMount -> su', mcc, su);
      this.get_user_allData(su).then(ud => {
        console.log('%c : componentDidMount -> ud', mcc, ud);
        // console.log('%c : componentDidMount -> this.org_depth', mcc, this.org_depth);
        this.build_hierarchy(ud, this.org_depth).then(h => {
          console.log('%c : componentDidMount -> hierarchy', mcc, h);
          this.setState({
            hierarchy_data: h[0],
            all_data: ud
          });
        });
      });
    });
  }

  public componentDidUpdate(prevProps: ISharePointOrgChartProps, prevState: ISharePointOrgChartState) {
    console.log('%c : componentDidUpdate -> prevState', mcc2, prevState);
    console.log('%c : componentDidUpdate -> this.state', mcc, this.state);
    // if (prevState.hierarchy_data !== this.state.hierarchy_data) {

    // }
  }

  public get_siteUsers = () => new Promise(resolve => {
    sp.web.siteUsers.get().then(su => {
      resolve(su);
    });
  })

  public async get_user_allData(su) {
    const promises = su.map((sui: any) => {
      const uname = sui.LoginName;
      return sp.profiles.getPropertiesFor(uname).then(ap => {
        // console.log('%c : ap', mcc, ap);
        if (ap && (
          (ap.ExtendedManagers && ap.ExtendedManagers.length)
          || (ap.ExtendedReports && ap.ExtendedReports.length > 1)
        )) {
          if (ap.ExtendedManagers.length > this.org_depth)
            this.org_depth = ap.ExtendedManagers.length;
          return (ap);
        }
        return null;
      });
    });
    const users = await Promise.all(promises);
    return (users.filter(u => u));
  }

  public async build_hierarchy(users, depth) {

    let h = {
      id: null,
      name: null,
      title: null,
      children: []
    };

    for (let i = 0; i < depth + 1; i++) {
      const level_i = users.filter(u => u.ExtendedManagers.length === i);
      // console.log('%c : build_hierarchy -> level_' + i, mcc, level_i);
      if (i === 0) {
        h.id = level_i[0].AccountName;
        h.name = level_i[0].DisplayName;
        h.title = level_i[0].Title;
      }
      else if (i === 1) {
        h.children =
          // h.children.push(
          level_i.map(l => {
            return ({
              id: l.AccountName,
              name: l.DisplayName,
              title: l.Title,
              children: []
            });
          }/* ) */
          );
      }
      else if (i === 2) {
        level_i.map(l => {
          // console.log('%c : build_hierarchy -> level 2 thing', mcc, l);
          const targ = h.children.filter(c => c.id === l.ExtendedManagers[1])[0];
          // console.log('%c : build_hierarchy -> targ', mcc, targ);
          targ.children.push(
            {
              id: l.AccountName,
              name: l.DisplayName,
              title: l.Title,
              children: []
            }
          );

        });
      }
      else if (i === 3) {
        const promises = level_i.map(l => {
          // console.log('%c : build_hierarchy -> level 3 thing', mcc, l);
          const gp = h.children.filter(c => c.id === l.ExtendedManagers[1])[0];
          // console.log('%c : build_hierarchy -> gp', mcc, gp);
          const targ = gp.children.filter(gc => gc.id === l.ExtendedManagers[2])[0];
          // console.log('%c : build_hierarchy -> targ', mcc, targ);
          targ.children.push(
            {
              id: l.AccountName,
              name: l.DisplayName,
              title: l.Title,
              children: []
            }
          );
          return h;
        });
        const hierarchy = await Promise.all(promises);
        return hierarchy;
      }
      // if (i == depth)
      //   console.log('%c : build_hierarchy -> h', mcc, h);
    }

  }

  public handler_card(id) {
    console.log('%c : handler_card -> id', mcc, id);
    const panel_data = this.state.all_data.filter(a => a.AccountName == id)[0];
    this.setState({
      panel_data: panel_data
    });
  }

  public handler_panel() {
    this.setState({ panel_data: null });
  }

  /*   public get_currentUser = () => new Promise(resolve => {
      sp.web.currentUser.get().then(cu => {
        // console.log('%c : SharePointOrgChart -> cu', mcc, cu);
        sp.web.siteUsers.getById(cu.Id).get().then(u => {
          resolve(u);
        });
      });
    })
  
    public get_currentUser_allProps = () => new Promise(resolve => {
      // sp.profiles.myProperties.get().then(ap => {
      sp.profiles.getPropertiesFor("i:0#.f|membership|kevinp@outwestrg.com").then(ap => {
        resolve(ap);
      });
    }) */

  public render(): React.ReactElement<ISharePointOrgChartProps> {
    console.log('%c : SharePointOrgChart -> this.state.hierarchy_data', mcc, this.state.hierarchy_data);
    return (
      <>
        <div dangerouslySetInnerHTML={{ __html: this.state.hierarchy_data }} />

      </>
    );
  }
}

/* <OrgChart
  hierarchy_data={this.state.hierarchy_data}
  handler={this.handler_card}
/> */
