import React from 'react';
import axios from "axios";
import { useState, useMemo, useCallback } from "react";
import { Header } from './components/Header';
import { Query } from './components/Query';
import { ExpenseItemList } from './components/ExpenseItemList';
import { ExpenseItem } from './components/ExpenseItem';
import { useWindowSize } from "./hooks/useWindowSize";
import configinfo from './serverconfig.json';

export const App = () => {

  const [expenseItemList, setExpenseItemList] = useState<any>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isError, setIsError] = useState(false);
  const [response, setResponse] = useState<any>("");
  const [errortext, setErrorText] = useState<any>("");
  
  const ServerAddress = configinfo.ServerAddress;
  const ServerPort = configinfo.ServerPort;
  const Username = configinfo.Username;
  const Password = configinfo.Password;
  const ApplicationName = configinfo.ApplicationName;
  const Protocol = configinfo.Protocol;
  
  const onClickFetchExpenseItemList = (keyword: any) => {
	
	setIsLoading(true);
    setIsError(false);
  
	axios
	  .get<any>(`${Protocol}://${ServerAddress}:${ServerPort}${ApplicationName}/SearchExpenseItem/z${keyword}?IRISUsername=${Username}&IRISPassword=${Password}`)
	  .then((result: any) => {
	  const eitems = result.data.map((eitem: any) => ({
		id: eitem.id,
		description: eitem.description
      }));
      setExpenseItemList(eitems);
	  })
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
      .finally(() => setIsLoading(false));
  };
  
   const onClickItem = useCallback((eitemid: any) => {
	setIsLoading(true);
	setIsError(false);

	axios
	   // eslint-disable-next-line
	  .get<any>(`${Protocol}://${ServerAddress}:${ServerPort}${ApplicationName}/ExpenseItemGetById/${eitemid}?IRISUsername=${Username}&IRISPassword=${Password}`)
	  .then((result: any) => {
		console.dir(result.data)
	    setResponse(result.data);
	  })
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
  // eslint-disable-next-line
  }, []);
  
  const [,height] = useWindowSize();
  
    // eslint-disable-next-line
    const ExpenseItemListMemo = useMemo(() => <ExpenseItemList isLoading = {isLoading} expenseItemList = {expenseItemList} onClickItem = {onClickItem} />, [onClickItem,expenseItemList]);

    return (
    <>
    <div className="title">
	<Header />
	</div>
    <div className="query">
	<Query onClickFetchExpenseItemList = {onClickFetchExpenseItemList} />
	{isError && <p style={{ color: "red" }}>エラーが発生しました　{`${errortext}`}</p>}
	</div>
    <div className="expenselist" style = {{ float: "left",width: "50%",height: `${height*0.9}px`,overflow: "auto",border: "solid #000000 1px"}}>	
    {ExpenseItemListMemo}
    </div>
    <div id="expensecontent" style = {{ width: "50%",height: `${height*0.9}px`,overflow: "auto",border: "solid #000000 1px"}}>
    <ExpenseItem response = {response} />
    </div>
    </>	
  );	
}
export default App;
