import * as React from 'react';
import styles from './Wrapper.module.scss';
import IWrapperProps from './IWrapperProps';
import SwaggerUI from "swagger-ui-react";

export default class Wrapper extends React.Component<IWrapperProps, {}> {
  public render(): React.ReactElement<IWrapperProps> {
    return (
      <div className={(this.props.isTeams) ? styles.teamsContext : styles.sharepointContext }>
        <SwaggerUI url={this.props.url} />
      </div>
    );
  }
}
