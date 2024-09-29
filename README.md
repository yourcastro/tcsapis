import React, { useState, useEffect } from 'react';
import axios from 'axios';
import loadingGif from './assets/loading.gif'; // Import the GIF

const ApiComponent = () => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true); // Loading state

  useEffect(() => {
    // Simulate API call
    const fetchData = async () => {
      setLoading(true); // Show loading GIF
      try {
        const response = await axios.get('https://jsonplaceholder.typicode.com/posts');
        setData(response.data);
      } catch (error) {
        console.error('Error fetching data:', error);
      } finally {
        setLoading(false); // Hide loading GIF after API response
      }
    };

    fetchData();
  }, []);

  return (
    <div>
      {loading ? (
        <div className="loading-container">
          <img src={loadingGif} alt="Loading..." className="loading-gif" />
        </div>
      ) : (
        <div className="data-container">
          {data ? (
            data.map((item) => <p key={item.id}>{item.title}</p>)
          ) : (
            <p>No data available</p>
          )}
        </div>
      )}
    </div>
  );
};

export default ApiComponent;



/* Center the loading GIF */
.loading-container {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100vh; /* Full screen height */
}

.loading-gif {
  width: 50px;
  height: 50px;
}

/* Style for the data container */
.data-container {
  padding: 20px;
}

Subject: Request for Praveen to Work Remotely This Week

Hi [Boss's Name],

I hope this message finds you well. I would like to request that Praveen work from home for the entire week. He has a significant amount of API integration and functionality work to complete for the upcoming UAT release. His current travel schedule is causing some delays, and being able to focus without the commute would greatly enhance his productivity.

Thank you for considering this request.

Best regards,
[Your Name]



