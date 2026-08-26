import "./LineChart.css";

export default function LineChart() {
  return (
    <div className="line-chart-container">

      {/* Y Axis */}
      <span className="line-y y220">220</span>
      <span className="line-y y165">165</span>
      <span className="line-y y110">110</span>
      <span className="line-y y55">55</span>
      <span className="line-y y0">0</span>

      {/* Grid */}
      <div className="line-grid grid220" />
      <div className="line-grid grid165" />
      <div className="line-grid grid110" />
      <div className="line-grid grid55" />
      <div className="line-grid grid0" />

      {/* Graph */}
      <svg
        className="line-graph"
        viewBox="0 0 1000 220"
        preserveAspectRatio="none"
      >
        {/* GREEN LINE */}
        <path
          className="graph-line green-line"
          d="
            M 80 135
            C 150 90, 220 75, 280 90
            C 350 105, 420 120, 500 110
            C 570 100, 620 25, 720 35
            C 800 40, 860 55, 920 65
          "
        />

        {/* GREEN POINTS */}
        <circle className="green-point" cx="80" cy="135" r="6" />
        <circle className="green-point" cx="280" cy="90" r="6" />
        <circle className="green-point" cx="500" cy="110" r="6" />
        <circle className="green-point" cx="720" cy="35" r="6" />
        <circle className="green-point" cx="920" cy="65" r="6" />

        {/* BLUE LINE */}
        <path
          className="graph-line blue-line"
          d="
            M 80 170
            C 150 120, 220 120, 280 130
            C 350 140, 420 155, 500 150
            C 570 140, 630 70, 720 88
            C 800 90, 860 105, 920 115
          "
        />

        {/* BLUE POINTS */}
        <circle className="blue-point" cx="80" cy="170" r="6" />
        <circle className="blue-point" cx="280" cy="130" r="6" />
        <circle className="blue-point" cx="500" cy="150" r="6" />
        <circle className="blue-point" cx="720" cy="88" r="6" />
        <circle className="blue-point" cx="920" cy="115" r="6" />
      </svg>

      {/* X Axis */}
      <div className="line-x-axis">
        <span>Mon</span>
        <span>Tue</span>
        <span>Wed</span>
        <span>Thu</span>
        <span>Fri</span>
      </div>

    </div>
  );
}