import * as React from "react";
import styles from "./Users.module.scss";
import type { IUsersProps } from "./IUsersProps";
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";

export default class Users extends React.Component<
  IUsersProps,
  { userProfile: any; email: string }
> {
  constructor(props: IUsersProps) {
    super(props);
    this.state = {
      userProfile: null,
      email: "",
    };
  }

  // Fetch current logged-in user's profile data
  async getUserProfile() {
    try {
      const graph = graphfi().using(SPFx(this.props.context));
      const currentUser = await graph.me();
      this.setState({ userProfile: currentUser });
    } catch (error) {
      console.error("Error fetching current user's profile:", error);
    }
  } // Fetch user data by email address
  async getUserProfileByEmail() {
    const { email } = this.state;
    if (!email) {
      alert("Please enter an email address");
      return;
    }
    try {
      const graph = graphfi().using(SPFx(this.props.context));
      const userByEmail = await graph.users.getById(email)();
      this.setState({ userProfile: userByEmail });
    } catch (error) {
      console.error(`Error fetching user profile for email ${email}:`, error);
    }
  }

  // Handle email input change
  handleEmailChange(event: React.ChangeEvent<HTMLInputElement>) {
    this.setState({ email: event.target.value });
  }
  public render(): React.ReactElement<IUsersProps> {
    const { userProfile, email } = this.state;
    return (
      <section className={styles.users}>
        <h2>Fetch User Profile</h2>

        <br />
        <div>
          <input
            type="email"
            value={email}
            onChange={(e) => this.handleEmailChange(e)}
            placeholder="Enter user email"
          />
          <button onClick={() => this.getUserProfileByEmail()}>
            Get User by Email
          </button>
        </div>
        {/* Button to fetch current logged-in user data */}
        <div>
          <button onClick={() => this.getUserProfile()}>
            Get Current User
          </button>
        </div>
        {/* Display user profile data */}
        {userProfile ? (
          <div>
            <h3>{userProfile.displayName}</h3>
            <p>Email: {userProfile.mail}</p>
            <p>Job Title: {userProfile.jobTitle}</p>
            <p>Phone: {userProfile.mobilePhone}</p>
            <p>Office Location: {userProfile.officeLocation}</p>
          </div>
        ) : (
          <p>No user profile data available.</p>
        )}
      </section>
    );
  }
}
