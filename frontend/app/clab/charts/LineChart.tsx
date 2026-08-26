import "./LineChart.css";

export default function LineChart() {
  return (
    <div className="clab-line-chart">
      {/* Y AXIS */}

      <span className="line-y-label y-220">220</span>
      <span className="line-y-label y-165">165</span>
      <span className="line-y-label y-110">110</span>
      <span className="line-y-label y-55">55</span>
      <span className="line-y-label y-0">0</span>

      {/* GRID */}

      <div className="line-grid line-grid-220" />
      <div className="line-grid line-grid-165" />
      <div className="line-grid line-grid-110" />
      <div className="line-grid line-grid-55" />
      <div className="line-grid line-grid-0" />

      {/* GRAPH */}

      <svg
        className="clab-line-svg"
        viewBox="0 0 1000 250"
        preserveAspectRatio="none"
      >
        {/* GREEN */}

        <path
          className="graph-path green-path"
          d="
            M 80 145
            C 150 95, 220 80, 280 90
            C 350 100, 420 120, 500 110
            C 580 100, 630 30, 720 35
            C 800 40, 860 55, 920 65
          "
        />

        <circle className="green-point" cx="80" cy="145" r="5" />
        <circle className="green-point" cx="280" cy="90" r="5" />
        <circle className="green-point" cx="500" cy="110" r="5" />
        <circle className="green-point" cx="720" cy="35" r="5" />
        <circle className="green-point" cx="920" cy="65" r="5" />

        {/* BLUE */}

        <path
          className="graph-path blue-path"
          d="
            M 80 185
            C 150 135, 220 125, 280 135
            C 350 145, 420 160, 500 155
            C 580 145, 630 85, 720 90
            C 800 92, 860 105, 920 120
          "
        />

        <circle className="blue-point" cx="80" cy="185" r="5" />
        <circle className="blue-point" cx="280" cy="135" r="5" />
        <circle className="blue-point" cx="500" cy="155" r="5" />
        <circle className="blue-point" cx="720" cy="90" r="5" />
        <circle className="blue-point" cx="920" cy="120" r="5" />
      </svg>

      {/* X AXIS */}

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