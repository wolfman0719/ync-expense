import React from 'react';
import { ChangeEvent, useState } from "react";

export const Query = (props: any) => {

  const {onClickFetchExpenseItemList} = props;
  const [inputtext, setInputText] = useState<any>("");
    
  const onChangeText = (e: ChangeEvent<HTMLInputElement>) => setInputText(e.target.value);
    
  return (
    <>
	  <label className="p-2">費用項目: </label>
	  <input type="text" value = {inputtext} onChange={onChangeText} />
	  <button className="btn btn-secondary" onClick={() => onClickFetchExpenseItemList(inputtext)}>費用項目検索</button>
    </>	
  );	
}
export default Query;
