import "./PieChart.css";

export default function PieChart() {
  return (
    <div className="pie-chart-container">
      <div className="pie-chart">
        <div className="pie-hole"></div>
      </div>

      <span className="pie-label label-tue">Tue:164</span>
      <span className="pie-label label-mon">Mon:118</span>
      <span className="pie-label label-fri">Fri:188</span>
      <span className="pie-label label-thu">Thu:201</span>
      <span className="pie-label label-wed">Wed:142</span>
    </div>
  );
}