import React from 'react';

export const ExpenseItemList = (props: any) => {

  const { isLoading, onClickItem, expenseItemList } = props;

  return (
    <>
      <ul className="collection" style={{ margin: 0 }}>
        {isLoading ? (
          <li className="collection-item">Data Loading</li>
        ) : (
          expenseItemList.map((eitem: any, index: number) => (
            <li key={index} className="collection-item" style={{ padding: 0 }}>
              <button
                className="btn-flat waves-effect"
                style={{ width: "100%", textAlign: "left", display: "flex", justifyContent: "space-between", alignItems: "center", padding: "0 16px" }}
                onClick={() => onClickItem(eitem.id)}
              >
                <span>{`${eitem.description}`}</span>
                <i className="material-icons">chevron_right</i>
              </button>
            </li>
          ))
        )}
      </ul>
    </>
  );
}
export default ExpenseItemList;
