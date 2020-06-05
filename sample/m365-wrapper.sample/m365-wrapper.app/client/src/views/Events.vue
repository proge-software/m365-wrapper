<template>
  <div class="about">
    <h1>This is an about page by Tacconi</h1>
  </div>
</template>

<script>
// @ is an alias to /src
//import HelloWorld from '@/components/HelloWorld.vue'

export default {
  name: "Events",
  components: {
    //HelloWorld
  },
  mounted() {
    const recaptchaScript = document.createElement("script");
    recaptchaScript.setAttribute(
      "src",
      "https://unpkg.com/m365-wrapper/_bundles/m365-wrapper.js"
    );
    document.head.appendChild(recaptchaScript);
    this.EntryPoint();
  },
  methods: {
    async EntryPoint(){

      const z = new window.M365Wrapper("9f43a6bd-9b42-4cf9-82f8-d9f1960596cc", "https://login.microsoftonline.com/common");
            
            const res02 = await z.StatLoginPopupProcess(); //OK
     }
    }
};
</script>
