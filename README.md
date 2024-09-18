export const HomefetchData = (data) => {
  console.log('ActionData', data)
  return async (dispatch, getState) => {
    var myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");
 
var requestOptions = {
  method: 'GET',
  headers: myHeaders,
  redirect: 'follow',
  mode:'no-cors',
  credentials: 'include', // Use this if you're sending cookies or need credentials,
};
 
fetch("https://159.208.208.142:5001/webservice/api/Home/GetAllWelcomePage", requestOptions)
  .then(response => response.text())
  .then(result => console.log('result',result))
  .catch(error => console.log('error', error));
 
 
    // dispatch(fetchDataRequest());
    try {
      //   const response = await axios.get('https://api.example.com/data');
      // dispatch(fetchDataSuccess(response));
    } catch (error) {
      // dispatch(fetchDataFailure(error.message));
    }
  };
};
