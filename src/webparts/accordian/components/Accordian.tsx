import * as React from 'react';
import styles from './Accordian.module.scss';
import type { IAccordianProps } from './IAccordianProps';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
 
import {Acc} from './Acc'
import { DisplayMode } from '@microsoft/sp-core-library';
import {addToList} from './service.js'



interface IAccordianState {
  isFormVisible: boolean;
  newItemTitle: string;
  newItemDescription: string;
}

export default class Accordian extends React.Component<IAccordianProps, IAccordianState> {
  constructor(props: IAccordianProps) {
    super(props);
    this.state = {
      isFormVisible: false,
      newItemTitle: '',
      newItemDescription: ''
    };
  } 


  private handleAddItemClick = (): void => {
    this.setState({ isFormVisible: true });
  };

  private handleDescriptionChange = (value: string): string => {
    this.setState({ newItemDescription: value });
    console.log(this.state.newItemDescription)
    return value; // Return the value to satisfy the RichText component's requirements
  };

  private handleInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const { name, value } = e.target;
    this.setState({ [name]: value } as unknown as Pick<IAccordianState, keyof IAccordianState>);
  };

  private  handleSubmit = async () => {
    const { newItemTitle, newItemDescription } = this.state;


    const res = await addToList(this.props.context,this.props.list,newItemTitle,newItemDescription)
    console.log(res);

    this.setState({ isFormVisible: false, newItemTitle: '', newItemDescription: '' });
  };

 

  public render(): React.ReactElement<IAccordianProps> {
    const { mode } = this.props;
    const { isFormVisible, newItemTitle, newItemDescription } = this.state;
  
    return (
      <div>
        {mode === DisplayMode.Edit && (
          <div>
            <button className={styles.button} onClick={this.handleAddItemClick}>Add Item</button>
            {isFormVisible && (
              <div className={styles.form}>
                <h2>Add New Item</h2>
                <label>
                  Title:
                  <input
                    type="text" 
                    name="newItemTitle"
                    value={newItemTitle}
                    onChange={this.handleInputChange}
                  />
                </label>
                <label>
                  Description:
                  <RichText
                    value={newItemDescription}
                    onChange={this.handleDescriptionChange}
                    isEditMode={true}
                 />
                </label>
                <button className={styles.button} onClick={this.handleSubmit}>Submit</button>
              </div>
            )}
          </div>
        )}

        
          
        {mode === DisplayMode.Read && (<Acc list={this.props.list} context={this.props.context}></Acc>)}
      </div>
    );
  }
}
