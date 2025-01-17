import React, { ChangeEvent,useState,useEffect } from 'react';
import axios from "axios";
import "./ExpenseItem.css"
import configinfo from '../serverconfig.json';

export const ExpenseItem = (props: any) => {

  const {response} = props;

  const [isLoading, setIsLoading] = useState(false);
  const [isError, setIsError] = useState(false);
  const [errorText, setErrorText] = useState("");
  const [description, setDescription] = useState("");
  const [paymentto, setPaymentto] = useState("");
  const [accounts, setAccounts] = useState("");
  const [amount, setAmount] = useState("");
  const [onbehalf, setOnBehalf] = useState("");
  const [deleted, setDeleted] = useState(false);

  useEffect( () => {

    setDescription(response.description)         
    setPaymentto(response.paymentto)         
    setAccounts(response.accounts)         
    setAmount(response.amount)         
    setOnBehalf(response.onbehalf)
    setDeleted(false)         
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [response]);   
  
  const serverAddress = configinfo.ServerAddress;
  const serverPort = configinfo.ServerPort;
  const username = configinfo.Username;
  const password = configinfo.Password;
  const applicationName = configinfo.ApplicationName;
  const protocol = configinfo.Protocol;

  const onChangeDescription = (e: ChangeEvent<HTMLInputElement>) => setDescription(e.target.value);
  const onChangePaymentto = (e: ChangeEvent<HTMLInputElement>) => setPaymentto(e.target.value);
  const onChangeAccounts = (e: ChangeEvent<HTMLInputElement>) => setAccounts(e.target.value);
  const onChangeAmount = (e: ChangeEvent<HTMLInputElement>) => setAmount(e.target.value);
  const onChangeBehalf = (e: ChangeEvent<HTMLInputElement>) => setOnBehalf(e.target.value); 
  
  const saveExpense = (e: any) => {

    setIsLoading(true);
    setIsError(false);
    
      const senddata: any =  {};
      senddata.new = 0;
      senddata.id = response.id;
      senddata.description = description;
      senddata.paymentto = paymentto;
      senddata.accounts = accounts;
      senddata.amount = amount;
      senddata.onbehalf = onbehalf;
  
      axios
        .post<any>(`${protocol}://${serverAddress}:${serverPort}${applicationName}/SaveExpenseItem?IRISUsername=${username}&IRISPassword=${password}`,senddata)
       .then((result: any) => {
           setIsError(false)
           alert('保存しました!!');})
          .catch((error: any) => {
           setIsError(true)
       if (error.response) {			
         setErrorText(error.response.data.summary);
       }
       else if (error.request) {
         setErrorText(error.request);
       } 
       else {
         setErrorText(error.message);
       }
        })
          .finally(() => setIsLoading(false))
  };
  
  const deleteExpense = (e: any) => {

    setIsLoading(true);
    setIsError(false);
    
      axios
        .get<any>(`${protocol}://${serverAddress}:${serverPort}${applicationName}/DeleteExpenseItemById/${response.id}?IRISUsername=${username}&IRISPassword=${password}`)
       .then((result: any) => {
           setIsError(false)
           setDeleted(true)
           setDescription("")
           setPaymentto("")
           setAccounts("")
           setAmount("")
           setOnBehalf("")
           alert('削除しました!!');})
          .catch((error: any) => {
           setIsError(true)
       if (error.response) {			
         setErrorText(error.response.data.summary);
       }
       else if (error.request) {
         setErrorText(error.request);
       } 
       else {
         setErrorText(error.message);
       }
        })
          .finally(() => setIsLoading(false))
  };
  
  return (
    <div className="container">
      {isError && <p style={{ color: "red" }}>エラーが発生しました　{`${errorText}`}</p>}
      {isLoading && <p>Loading...</p>}
      <b>経費項目</b>
      <hr />
      <table>
      <div>
        <tr>
        <td width={100}><label>備考：</label></td>
        <td><input type="text"  value={description}  onChange={onChangeDescription} size={50} style = {{float: "left"}}/></td>
        </tr>
      </div>
      <div className="spacer" />
      <div>
        <tr>
        <td width={100}><label>支払先：</label></td>
        <td><input type="text" value={paymentto} onChange={onChangePaymentto} style = {{float: "left"}}/></td>
        </tr>
      </div>
      <div className="spacer" />
      <div>
        <tr>
        <td width={100}><label>勘定科目：</label></td>
        <td><input type="text" value={accounts} onChange={onChangeAccounts} style = {{float: "left"}}/></td>
        </tr>
      </div>
      <div className="spacer" />
      <div>
        <tr>
        <td width={100}><label>金額：</label></td>
        <td><input type="text" value={amount} onChange={onChangeAmount} style = {{float: "left"}}/></td>
        </tr>
      </div>
      <div className="spacer" />
      <div>
        <tr>
        <td width={100}><label>立て替え：</label></td>
        <td><input type="text" value={onbehalf} onChange={onChangeBehalf} style = {{float: "left"}}/></td>
        </tr>
      </div>
      </table>
      <div className="spacer" />
      {(deleted == false) && <button className = "btn btn-outline-primary" style = {{textAlign: "left"}} onClick={saveExpense}>保存</button>}
      {(deleted == false) && <button className = "btn btn-outline-danger" style = {{textAlign: "left"}} onClick={deleteExpense}>削除</button>}
    </div>  );	
}
export default ExpenseItem;
