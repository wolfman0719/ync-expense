import React from 'react';
import { ChangeEvent, useState } from "react";

export const Query = (props: any) => {

  const { onClickFetchExpenseItemList } = props;
  const [inputtext, setInputText] = useState<any>("");

  const onChangeText = (e: ChangeEvent<HTMLInputElement>) => setInputText(e.target.value);

  const onKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Enter") onClickFetchExpenseItemList(inputtext);
  };

  return (
    <div style={{ display: "inline-flex", alignItems: "center", border: "1px solid #bbb", borderRadius: "24px", padding: "4px 12px", background: "#fff" }}>
      <input
        id="expense-query"
        type="text"
        placeholder="費用項目"
        value={inputtext}
        onChange={onChangeText}
        onKeyDown={onKeyDown}
        style={{ border: "none", outline: "none", boxShadow: "none", margin: 0, padding: 0, fontSize: "1rem", width: "200px" }}
      />
      <i
        className="material-icons"
        style={{ cursor: "pointer", color: "#555", fontSize: "1.25rem", userSelect: "none" }}
        onClick={() => onClickFetchExpenseItemList(inputtext)}
      >search</i>
    </div>
  );
}
export default Query;
