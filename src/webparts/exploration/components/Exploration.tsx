import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IExplorationProps } from './IExplorationProps';
import styles from './Exploration.module.scss'; 
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
export default class Exploration extends React.Component<IExplorationProps, any> {
//SESSION///
  constructor(props: IExplorationProps) {
    super(props);

    this.state = {
      title: '',
      count: '',
      comments: '',
      date: '',
      technology: '',
      choices: [], 
      employees: [] as number[], 
      pickerKey: 0
    };
  }

  // ✅ Load choices (dynamic)
  public componentDidMount(): void {

    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SpfxFormDetails')/fields/getbyinternalnameortitle('Technolgy')`;

    this.props.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    )
    .then(res => res.json())
    .then(data => {
      this.setState({
        choices: data.Choices
      });
    });
  }

  public render(): React.ReactElement<IExplorationProps> {

    // ✅ SAVE FUNCTION
    const saveData = () => {

      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SpfxFormDetails')/items`;

      this.props.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify({
            Title: this.state.title,
            Count: Number(this.state.count),
            Comments: this.state.comments,
            Technolgy: this.state.technology,
            
 EmployeesId: 
this.state.employees[0]
  ,
            Clock: new Date(this.state.date).toISOString()
            
          })
        }
      )
      .then((res: SPHttpClientResponse) => {
        if (res.ok) {
if (!isFormValid) {
  alert("Please fill all fields");
  return;
}
          alert("Saved successfully ✅");
          console.log("Employees:", this.state.employees);

          // ✅ RESET FORM AFTER SAVE
          this.setState({
            title: '',
            count: '',
            comments: '',
            date: '',
            technology: '', 
             employees: [], 
        pickerKey: this.state.pickerKey + 1
          });

        } else {
          alert("Error saving ");
          console.log("Employees:", this.state.employees);
        }
      });
    };
const isFormValid =
  this.state.title &&
  this.state.count &&
  this.state.comments &&
  this.state.date &&
  this.state.technology &&
  this.state.employees.length > 0;
    // ✅ CANCEL FUNCTION
    const cancelForm = () => {
      this.setState({
        title: '',
        count: '',
        comments: '',
        date: '',
        technology: '' , 
        employees: [], 
        pickerKey: this.state.pickerKey + 1
      });
    };

  return (
  <div className={styles.pageContainer}>

    <div className={styles.formContainer}>

      <h3 className={styles.title}>Form</h3>

      <input
        className={styles.inputField}
        type="text"
        placeholder="Enter Title"
        value={this.state.title}
        onChange={(e) => this.setState({ title: e.target.value })}
      />

      <select
        className={styles.selectField}
        value={this.state.technology}
        onChange={(e) => this.setState({ technology: e.target.value })}
      >
        <option value="">Select Technology</option>
        {this.state.choices?.map((choice: string) => (
          <option key={choice} value={choice}>
            {choice}
          </option>
        ))}
      </select>

      <input
        className={styles.inputField}
        type="number"
        placeholder="Enter Count"
        value={this.state.count}
        onChange={(e) => this.setState({ count: e.target.value })}
      />
///hi
      <textarea
        className={styles.inputField}
        placeholder="Enter Comments"
        value={this.state.comments}
        onChange={(e) => this.setState({ comments: e.target.value })}
      />

      <input
        className={styles.inputField}
        type="datetime-local"
        value={this.state.date}
        onChange={(e) => this.setState({ date: e.target.value })}
      />
 <PeoplePicker
  key={this.state.pickerKey}  
  context={this.props.context as any}
  titleText="Select User(s)"
  personSelectionLimit={3}
  showtooltip={true}
  principalTypes={[PrincipalType.User]}
  resolveDelay={1000}
  ensureUser={true}
  webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
onChange={(items: any[]) => {

  

  const ids = items.map(i => i.id);

  this.setState({
    employees: ids
  });

}}
/>
    <div className={styles.buttonGroup}> 
  <div style={{ marginBottom: "8px" }}>
  {!isFormValid && (
    <span style={{ color: "red", fontSize: "12px" }}>
      Fill all fields and select user
    </span>
  )}
</div>
        <button className={styles.primaryButton} onClick={saveData} disabled={!isFormValid} >Save</button>
        <button className={styles.secondaryButton} onClick={cancelForm}>Cancel</button>
      </div>

    </div>

  </div>
);

  }
}
