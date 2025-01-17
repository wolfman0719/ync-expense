import React from 'react';

export const ExpenseItemList = (props: any) => {

const {isLoading, onClickItem, expenseItemList} = props;
    
  return (
    <>
	<table style = {{width: "100%"}}><tbody>
	  {isLoading ? (<tr><td>Data Loading</td></tr>)
		 : (
		 expenseItemList.map((eitem: any, index: number) => (
		 <tr key={index}>
		 <td><button className = "btn btn-outline-primary" style = {{width: "100%", textAlign: "left"}} onClick={() => onClickItem(eitem.id)}>{`${eitem.description}`}<i className="bi bi-chevron-right float-end"></i></button></td>
		 </tr>
		 )))
	  }
	</tbody></table>
    </>	
  );	
}
export default ExpenseItemList;
