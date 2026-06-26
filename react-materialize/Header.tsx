import React from 'react';

export const Header = () => {
    
  return (
    <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
      <img src="../yncorporation.png" alt="YN Corporation LLC" style={{ height: "4rem", width: "auto", border: "0" }} />
      <p className="flow-text blue-text text-darken-2" style={{ fontSize: "2rem", fontWeight: "bold", margin: 0 }}>経費項目編集</p>
    </div>
  );
}
export default Header;
